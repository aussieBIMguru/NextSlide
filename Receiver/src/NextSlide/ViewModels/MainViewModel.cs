using System;
using System.Collections.ObjectModel;
using System.Linq;
using NextSlide.Models;
using NextSlide.Mvvm;
using NextSlide.Services;

namespace NextSlide.ViewModels;

/// <summary>
/// Backs MainWindow. Owns the Session/Sheet lock/release flow, the
/// PowerPoint presentation picker, the command log, and the
/// SlidePollingService that ties them together — see the project's
/// handover doc §5 for the receiver design this implements.
///
/// State machine: SessionName and SheetUrl are both free-text together
/// until <see cref="LockSessionCommand"/> locks them as one step (see
/// <see cref="IsSessionLocked"/> / <see cref="IsHookupEditable"/>) — Lock
/// only becomes available once both a non-empty Session name and a usable
/// SheetUrl are present, so the two are validated and committed together
/// rather than sequentially. Locking unlocks the presentation picker (see
/// <see cref="IsPresentationPickerEnabled"/>); polling only ever runs while
/// a presentation is selected on top of that — <see cref="UpdatePollingTargetsAndState"/>
/// is the single place that starts/stops it. Releasing stops polling and
/// clears the presentation list, but deliberately leaves both SessionName
/// and SheetUrl in place (editable again) so re-locking doesn't require
/// retyping them. <see cref="OnPresentationUnavailable"/> drops back to an
/// unselected presentation (without unlocking) whenever PowerPoint or the
/// target presentation disappears mid-session — see SlidePollingService's
/// PresentationUnavailable doc comment.
/// </summary>
public sealed class MainViewModel : ObservableObject, IDisposable
{
    private readonly AppSettings _settings;
    private readonly SettingsService _settingsService;
    private readonly CommandDedupeStore _dedupeStore;
    private readonly SlidePollingService _pollingService;
    private readonly PowerPointController _powerPoint = new();

    private string _sessionName;
    private bool _isSessionLocked;
    private string _sheetUrl;
    private ObservableCollection<PresentationOption> _availablePresentations = new();
    private PresentationOption? _selectedPresentation;
    private string _statusMessage = "Enter a Session name and Sheet URL, then click Lock to begin.";
    private bool _isPolling;

    public MainViewModel(AppSettings settings, SettingsService settingsService)
    {
        _settings = settings;
        _settingsService = settingsService;

        // Pre-fill from last run for convenience, but never auto-lock —
        // polling (and therefore driving PowerPoint) only ever starts from
        // an explicit user action.
        _sessionName = settings.LastSessionName ?? "";
        _sheetUrl = settings.LastSheetUrl ?? "";

        // In-memory only — see CommandDedupeStore's doc comment for why a
        // fresh, empty store on every launch is the safer default (no
        // stale entries piling up across sessions, no risk of re-firing).
        _dedupeStore = new CommandDedupeStore();

        _pollingService = new SlidePollingService(_dedupeStore);
        _pollingService.CommandProcessed += OnCommandProcessed;
        _pollingService.PollError += OnPollError;
        _pollingService.PresentationUnavailable += OnPresentationUnavailable;

        CommandLog = new ObservableCollection<CommandLogItemViewModel>();

        LockSessionCommand = new RelayCommand(LockSession,
            () => !IsSessionLocked && !string.IsNullOrWhiteSpace(SessionName) && IsSheetUrlUsable);
        ReleaseSessionCommand = new RelayCommand(ReleaseSession, () => IsSessionLocked);
        RefreshPresentationsCommand = new RelayCommand(RefreshPresentations, () => IsSessionLocked);
        ClearLogCommand = new RelayCommand(() => CommandLog.Clear());
        SaveSettingsNowCommand = new RelayCommand(SaveSettingsNow);
    }

    public string AppName => AppInfo.Name;

    public string Monogram => AppInfo.Monogram;

    public string SubHeaderText => "Watches a Google Sheet for remote-clicker commands and drives PowerPoint.";

    public string SessionName
    {
        get => _sessionName;
        set
        {
            if (SetProperty(ref _sessionName, value))
                LockSessionCommand.RaiseCanExecuteChanged();
        }
    }

    public bool IsSessionLocked
    {
        get => _isSessionLocked;
        private set
        {
            if (!SetProperty(ref _isSessionLocked, value))
                return;

            LockSessionCommand.RaiseCanExecuteChanged();
            ReleaseSessionCommand.RaiseCanExecuteChanged();
            RefreshPresentationsCommand.RaiseCanExecuteChanged();
            OnPropertyChanged(nameof(IsHookupEditable));
            OnPropertyChanged(nameof(IsPresentationPickerEnabled));
        }
    }

    /// <summary>Both SessionName's and SheetUrl's TextBoxes bind IsEnabled here — free-text together until Lock, read-only together once locked.</summary>
    public bool IsHookupEditable => !IsSessionLocked;

    public string SheetUrl
    {
        get => _sheetUrl;
        set
        {
            if (!SetProperty(ref _sheetUrl, value))
                return;

            OnPropertyChanged(nameof(IsSheetUrlUsable));
            OnPropertyChanged(nameof(IsPresentationPickerEnabled));
            LockSessionCommand.RaiseCanExecuteChanged();
            UpdatePollingTargetsAndState();
        }
    }

    /// <summary>True once SheetUrl looks like a real Google Sheets link — gates unlocking the presentation picker, per the UI spec.</summary>
    public bool IsSheetUrlUsable => GoogleSheetReader.TryParseSheetUrl(SheetUrl, out _, out _);

    public bool IsPresentationPickerEnabled => IsSessionLocked && IsSheetUrlUsable;

    public ObservableCollection<PresentationOption> AvailablePresentations
    {
        get => _availablePresentations;
        private set => SetProperty(ref _availablePresentations, value);
    }

    public PresentationOption? SelectedPresentation
    {
        get => _selectedPresentation;
        set
        {
            if (SetProperty(ref _selectedPresentation, value))
                UpdatePollingTargetsAndState();
        }
    }

    public string StatusMessage
    {
        get => _statusMessage;
        private set => SetProperty(ref _statusMessage, value);
    }

    /// <summary>Drives the status row's "live" lamp — true only while the poll loop is actually running.</summary>
    public bool IsPolling
    {
        get => _isPolling;
        private set => SetProperty(ref _isPolling, value);
    }

    /// <summary>Newest first — only ever holds rows processed since this app launched (see CommandLogItemViewModel).</summary>
    public ObservableCollection<CommandLogItemViewModel> CommandLog { get; }

    public RelayCommand LockSessionCommand { get; }
    public RelayCommand ReleaseSessionCommand { get; }
    public RelayCommand RefreshPresentationsCommand { get; }
    public RelayCommand ClearLogCommand { get; }
    public RelayCommand SaveSettingsNowCommand { get; }

    private void LockSession()
    {
        IsSessionLocked = true;

        // Both committed eagerly and together — locking is the one
        // explicit action that validates both fields at once, so both get
        // persisted here rather than SheetUrl waiting until a presentation
        // is also picked (see UpdatePollingTargetsAndState, which used to
        // be the only place LastSheetUrl was saved).
        _settings.LastSessionName = SessionName.Trim();
        _settings.LastSheetUrl = SheetUrl.Trim();
        _settingsService.Save(_settings);

        StatusMessage = $"Session '{SessionName.Trim()}' locked. Pick a presentation to continue.";
        RefreshPresentations();
    }

    private void ReleaseSession()
    {
        _pollingService.Stop();
        IsPolling = false;
        IsSessionLocked = false;
        AvailablePresentations = new ObservableCollection<PresentationOption>();
        SelectedPresentation = null;
        StatusMessage = "Released — session and sheet URL kept. Lock again to resume.";
    }

    /// <summary>
    /// Fired when SlidePollingService discovers, mid-poll, that the target
    /// presentation is gone (PowerPoint closed, or just that presentation
    /// closed) — the case that used to be able to crash the app before a
    /// command attempt was left unguarded. Drops back to an unselected
    /// presentation (stopping polling as a side effect, via
    /// UpdatePollingTargetsAndState) without unlocking the session, then
    /// re-scans so the picker reflects reality — either PowerPoint's fully
    /// gone (empty list, explanatory status) or another open presentation
    /// is available to pick straight away.
    /// </summary>
    private void OnPresentationUnavailable(object? sender, EventArgs e)
    {
        SelectedPresentation = null;
        RefreshPresentations();
    }

    /// <summary>
    /// Re-scans PowerPoint's currently open presentations. Manual (via the
    /// Refresh button, and once automatically on Lock) rather than polled
    /// on a timer of its own — SlidePollingService already owns the one
    /// recurring COM+network loop, and a second independent timer here
    /// would just be two things able to touch COM concurrently for no
    /// real benefit.
    /// </summary>
    private void RefreshPresentations()
    {
        if (!IsSessionLocked)
            return;

        var previouslySelectedName = SelectedPresentation?.Name;
        var options = _powerPoint.ListOpenPresentations(out var diagnostic);
        AvailablePresentations = new ObservableCollection<PresentationOption>(options);

        SelectedPresentation = previouslySelectedName is null
            ? null
            : AvailablePresentations.FirstOrDefault(p => p.Name == previouslySelectedName);

        if (AvailablePresentations.Count == 0)
            StatusMessage = diagnostic ?? "No open presentations found — open one in PowerPoint, then Refresh.";
    }

    /// <summary>
    /// The single place that starts or stops SlidePollingService, called
    /// from every property setter that could change whether all three
    /// preconditions (locked, usable Sheet URL, a chosen presentation) are
    /// currently met.
    /// </summary>
    private void UpdatePollingTargetsAndState()
    {
        if (IsSessionLocked && SelectedPresentation is not null &&
            GoogleSheetReader.TryParseSheetUrl(SheetUrl, out var sheetId, out var gid))
        {
            _settings.LastSheetUrl = SheetUrl.Trim();

            _pollingService.Session = SessionName.Trim();
            _pollingService.GvizUrl = GoogleSheetReader.BuildGvizUrl(sheetId, gid);
            _pollingService.TargetPresentationName = SelectedPresentation.Name;

            if (!_pollingService.IsRunning)
            {
                _pollingService.Start();
                IsPolling = true;
                StatusMessage = $"Watching session '{SessionName.Trim()}' — hooked to '{SelectedPresentation.Name}'.";
            }
        }
        else if (_pollingService.IsRunning)
        {
            _pollingService.Stop();
            IsPolling = false;
            StatusMessage = "Paused — pick an open presentation to resume.";
        }
    }

    private void OnCommandProcessed(object? sender, SlideCommandProcessedEventArgs e)
    {
        // SlidePollingService's DispatcherTimer.Tick already runs on the UI
        // thread, so this is safe to touch the bound collection directly.
        CommandLog.Insert(0, new CommandLogItemViewModel(e.Row, e.Outcome, e.Detail));
    }

    private void OnPollError(object? sender, string message) => StatusMessage = message;

    /// <summary>
    /// Copies live view-model state back into the AppSettings instance that
    /// will be serialized on exit. Called from MainWindow.PersistWindowState
    /// just before SettingsService.Save. LastSessionName/LastSheetUrl are
    /// already kept current as the user edits them (see LockSession and
    /// UpdatePollingTargetsAndState), so there's nothing further to copy here.
    /// </summary>
    public void PersistToSettings()
    {
    }

    private void SaveSettingsNow()
    {
        PersistToSettings();
        _settingsService.Save(_settings);
    }

    public void Dispose()
    {
        _pollingService.CommandProcessed -= OnCommandProcessed;
        _pollingService.PollError -= OnPollError;
        _pollingService.PresentationUnavailable -= OnPresentationUnavailable;
        _pollingService.Dispose();
    }
}

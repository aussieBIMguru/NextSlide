namespace NextSlide.Models;

/// <summary>
/// The three ways this app can start and behave when its main window is
/// closed. One mechanism (App.xaml.cs + MainWindow's Closing handling)
/// serves all three — switching modes is a settings/argument change, never
/// a rewrite. See README.md "Run Modes" for how to change the default or
/// force a mode from the command line.
/// </summary>
public enum RunMode
{
    /// <summary>No window on launch — tray icon only ("headless").</summary>
    Silent,

    /// <summary>Window shows on launch; closing it exits the app normally. Template default.</summary>
    WindowedExitOnClose,

    /// <summary>Window shows on launch; closing it hides to tray instead of exiting.</summary>
    WindowedTrayOnClose
}

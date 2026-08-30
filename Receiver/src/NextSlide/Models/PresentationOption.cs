namespace NextSlide.Models;

/// <summary>
/// One open presentation in the running PowerPoint instance, as offered by
/// <see cref="Services.PowerPointController.ListOpenPresentations"/> for the
/// "hook to" combobox. <see cref="Name"/> is PowerPoint's own Presentation
/// name (its window/file title, e.g. "Q3 Review.pptx") — the only handle
/// COM automation needs to re-find it on each poll tick, since a live COM
/// reference can't safely be cached across ticks (the user might close it).
/// </summary>
public sealed record PresentationOption(string Name, bool IsInSlideShow)
{
    /// <summary>What the ComboBox displays — flags whether it's actually presenting right now, since that's what gates whether a command can fire.</summary>
    public override string ToString() => IsInSlideShow ? $"{Name}  (presenting)" : $"{Name}  (not presenting)";
}

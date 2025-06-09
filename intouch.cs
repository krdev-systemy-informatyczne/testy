foreach (var style in styles.Elements<Style>())
{
    if (!expectedStyle.Select(s => s.StyleId).Contains(style.StyleId))
    {
        // Ukrywamy style, które nie są oczekiwane

        if (style.Elements<SemiHidden>().FirstOrDefault() == null)
            style.AppendChild(new SemiHidden());

        if (style.Elements<UnhideWhenUsed>().FirstOrDefault() == null)
            style.AppendChild(new UnhideWhenUsed());

        if (style.Elements<UiPriority>().FirstOrDefault() == null)
            style.AppendChild(new UiPriority() { Val = 99 });
        else
            style.Elements<UiPriority>().First().Val = 99;
    }
    else
    {
        // Jeśli styl jest oczekiwany — odblokuj
        var semiHidden = style.Elements<SemiHidden>().FirstOrDefault();
        if (semiHidden != null)
            semiHidden.Remove();

        var unhideWhenUsed = style.Elements<UnhideWhenUsed>().FirstOrDefault();
        if (unhideWhenUsed != null)
            unhideWhenUsed.Remove();

        var uiPriority = style.Elements<UiPriority>().FirstOrDefault();
        if (uiPriority != null)
            uiPriority.Val = 1;  // np. 1 dla oczekiwanych stylów — będzie wysoko w UI
    }
}

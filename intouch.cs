
// Ukryj pozostałe style
    foreach (var style in styles.Elements<Style>())
    {
        if (!expectedStyle.Select(s => s.StyleId).Contains(style.StyleId))
        {
            // Dodaj jeśli nie istnieje — nie zawsze style mają te elementy
            if (style.Elements<SemiHidden>().FirstOrDefault() == null)
                style.AppendChild(new SemiHidden());

            if (style.Elements<UnhideWhenUsed>().FirstOrDefault() == null)
                style.AppendChild(new UnhideWhenUsed());

            style.StyleUiPriority = new StyleUiPriority() { Val = 99 };
        }
        else
        {
            // Jeśli styl jest na liście oczekiwanych — upewnij się że NIE jest ukryty
            var semiHidden = style.Elements<SemiHidden>().FirstOrDefault();
            if (semiHidden != null)
                semiHidden.Remove();
        }
    }

    styles.Save();

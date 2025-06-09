public void FixMissingStylesForText(WordprocessingDocument doc)
{
    var expectedStyle = GuidebookQualitySettings.GetStylesForText();
    var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Brak głównej części dokumentu");

    var stylePart = mainPart.StyleDefinitionsPart;
    if (stylePart == null)
    {
        stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
    }

    if (stylePart.Styles == null)
    {
        stylePart.Styles = new Styles();
    }

    var styles = stylePart.Styles;

    // --- Dodanie brakujących stylów lub poprawienie istniejących ---
    foreach (var expected in expectedStyle)
    {
        var existing = styles.Elements<Style>().FirstOrDefault(s => s.StyleId == expected.StyleId);

        if (existing == null)
        {
            // Brak — dodaj styl
            styles.AppendChild(expected);
        }
        else
        {
            // Styl istnieje — odsłoń go

            // Usuń SemiHidden
            existing.RemoveAllChildren<SemiHidden>();

            // Dodaj PrimaryStyle jeśli brak
            if (existing.Elements<PrimaryStyle>().FirstOrDefault() == null)
            {
                existing.Append(new PrimaryStyle());
            }

            // Usuń UnhideWhenUsed
            existing.RemoveAllChildren<UnhideWhenUsed>();

            // UWAGA: czasem trzeba też podnieść UI priority, żeby Word pokazywał ten styl "wyżej"
            // Jeżeli masz PrimaryStyle, warto UI priority dać niskie (np. 1 lub 99)
            var uiPriority = existing.Elements<UIPriority>().FirstOrDefault();
            if (uiPriority == null)
            {
                existing.Append(new UIPriority() { Val = 1 });
            }
            else
            {
                uiPriority.Val = 1;
            }
        }
    }

    // --- Ukrycie wszystkich innych stylów ---
    foreach (var style in styles.Elements<Style>())
    {
        bool isExpected = expectedStyle.Any(s => s.StyleId == style.StyleId);

        if (!isExpected)
        {
            // Dodaj SemiHidden (jeśli brak)
            if (style.Elements<SemiHidden>().FirstOrDefault() == null)
            {
                style.Append(new SemiHidden());
            }

            // Usuń PrimaryStyle — żeby nie pokazywał się na pasku
            style.RemoveAllChildren<PrimaryStyle>();

            // Usuń UnhideWhenUsed
            style.RemoveAllChildren<UnhideWhenUsed>();
        }
    }

    styles.Save();
}

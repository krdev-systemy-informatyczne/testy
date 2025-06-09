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
    var existingStyleIds = styles.Elements<Style>()
        .Select(s => s.StyleId)
        .ToHashSet();

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

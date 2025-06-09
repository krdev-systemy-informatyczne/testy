public void FixMissingStylesForText(WordprocessingDocument doc)
{
    var expectedStyles = GuidebookQualitySettings.GetStylesForText();
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

    // Dodanie brakujących stylów lub poprawienie istniejących
    foreach (var expected in expectedStyles)
    {
        var existing = styles.Elements<Style>().FirstOrDefault(s => s.StyleId == expected.StyleId);

        if (existing == null)
        {
            styles.AppendChild(expected);
        }
        else
        {
            // Styl istnieje - uczyń go widocznym

            // Usuń SemiHidden
            existing.RemoveAllChildren<SemiHidden>();

            // Usuń UnhideWhenUsed
            existing.RemoveAllChildren<UnhideWhenUsed>();

            // Dodaj PrimaryStyle jeśli brak
            if (existing.Elements<PrimaryStyle>().FirstOrDefault() == null)
            {
                existing.Append(new PrimaryStyle());
            }

            // Ustaw niskie UI Priority (im niższe, tym wyżej na pasku, np. 1)
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

    // Ukrycie WSZYSTKICH INNYCH stylów
    foreach (var style in styles.Elements<Style>())
    {
        bool isExpected = expectedStyles.Any(s => s.StyleId == style.StyleId);

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

            // Opcjonalnie podnieś UI priority żeby był niżej
            var uiPriority = style.Elements<UIPriority>().FirstOrDefault();
            if (uiPriority == null)
            {
                style.Append(new UIPriority() { Val = 99 });
            }
            else
            {
                uiPriority.Val = 99;
            }
        }
    }

    styles.Save();
}

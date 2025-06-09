private void SetStyleVisibility(Style style, bool visible)
{
    style.RemoveAllChildren<PrimaryStyle>();
    if (visible)
    {
        style.Append(new PrimaryStyle());
    }

    style.RemoveAllChildren<SemiHidden>();
    if (!visible)
    {
        style.Append(new SemiHidden());
    }

    style.RemoveAllChildren<UnhideWhenUsed>();

    var uiPriority = style.Elements<UIPriority>().FirstOrDefault();
    if (uiPriority == null)
    {
        uiPriority = new UIPriority();
        style.Append(uiPriority);
    }
    uiPriority.Val = visible ? 1U : 99U;
}

private void ForceSetBasedOnToNormal(Style style)
{
    style.RemoveAllChildren<BasedOn>();
    style.BasedOn = new BasedOn() { Val = "Normal" };
}

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

    foreach (var expected in expectedStyles)
    {
        var existing = styles.Elements<Style>().FirstOrDefault(s => s.StyleId == expected.StyleId);

        if (existing == null)
        {
            styles.AppendChild(expected);
        }
        else
        {
            SetStyleVisibility(existing, visible: true);
        }
    }

    foreach (var style in styles.Elements<Style>())
    {
        bool isExpected = expectedStyles.Any(s => s.StyleId == style.StyleId);

        if (!isExpected)
        {
            SetStyleVisibility(style, visible: false);
            ForceSetBasedOnToNormal(style); // KLUCZOWE - bez tego Word nadal pokazuje np. Heading1
        }
    }

    styles.Save();
}

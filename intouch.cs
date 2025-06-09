public QualityResult CheckFooterQuality(WordprocessingDocument document)
{
    var issues = new List<QualityIssue>();
    var expectedTitle = document.PackageProperties.Title ?? string.Empty;
    expectedTitle = expectedTitle.Replace("picipolo", string.Empty).Trim();

    var mainPart = document.MainDocumentPart;
    if (mainPart == null)
    {
        issues.Add(new QualityIssue(QualityIssueCode.MissingMainPart, "Brak głównej części dokumentu."));
        return new QualityResult(nameof(CheckFooterQuality), EQualityTypeResult.Error, issues);
    }

    // Sprawdzenie istnienia Footer
    var footerPart = mainPart.FooterParts.FirstOrDefault();
    if (footerPart == null)
    {
        issues.Add(new QualityIssue(QualityIssueCode.FooterMissing, "Brak stopki w dokumencie."));
        return new QualityResult(nameof(CheckFooterQuality), EQualityTypeResult.Error, issues);
    }

    // Sprawdzenie wysokości stopki (SectionProperties.FooterMargin) - to nie jest idealna metoda
    var sectionProps = mainPart.Document.Body.Elements<SectionProperties>().FirstOrDefault();
    var pageMargin = sectionProps?.GetFirstChild<PageMargin>();
    if (pageMargin != null)
    {
        var footerMarginTwips = pageMargin.Footer?.Value ?? 0;
        const double twipsPerMm = 56.692913; // 1 mm = ~56.69 twips
        double footerMarginMm = footerMarginTwips / twipsPerMm;

        // Sprawdzamy tolerancję ±0.5 mm
        if (Math.Abs(footerMarginMm - 9.0) > 0.5)
        {
            issues.Add(new QualityIssue(QualityIssueCode.FooterHeightIncorrect,
                $"Aktualna wysokość marginesu stopki to {footerMarginMm:F1} mm, powinna wynosić ok. 9 mm."));
        }
    }

    // Sprawdzanie treści stopki
    string footerText = footerPart.Footer?.InnerText ?? string.Empty;

    // Sprawdzenie czy jest numeracja stron (wg slajdu: "Strona X z Y")
    if (string.IsNullOrWhiteSpace(footerText) ||
        (!footerText.Contains("Strona", StringComparison.OrdinalIgnoreCase) &&
         !footerText.Contains("Page", StringComparison.OrdinalIgnoreCase)))
    {
        issues.Add(new QualityIssue(QualityIssueCode.MissingPageNumber, "Brak numeracji stron w stopce."));
    }

    // Sprawdzenie czy występuje tytuł dokumentu w stopce (obok numeracji)
    if (!string.IsNullOrWhiteSpace(expectedTitle) &&
        !footerText.Contains(expectedTitle, StringComparison.OrdinalIgnoreCase))
    {
        issues.Add(new QualityIssue(QualityIssueCode.FooterMissingTitle,
            "Brak oczekiwanego tytułu dokumentu w stopce."));
    }
    
    // Sprawdzenie czy treść w stopce jest ukryta (dla czytników ekranów)
    var paragraphs = footerPart.Footer.Descendants<Paragraph>();
    bool anyVisibleContent = paragraphs.Any(p =>
        p.Descendants<Run>().Any(run =>
            run.RunProperties?.Vanish == null && !string.IsNullOrWhiteSpace(run.InnerText)));

    if (anyVisibleContent)
    {
        issues.Add(new QualityIssue(QualityIssueCode.FooterContentExposedToReader,
            "Stopka zawiera widoczny tekst, który może być odczytany przez czytnik ekranu. Jeśli to nie jest treść merytoryczna, oznacz ją jako ukrytą."));
    }

    return new QualityResult(nameof(CheckFooterQuality),
        issues.Any(i => i.Type == EQualityTypeResult.Error) ? EQualityTypeResult.Error : EQualityTypeResult.Ok,
        issues);
}

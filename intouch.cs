public List<Style> GetStylesForText()
{
    return new List<Style>()
    {
        // Nagłówki
        InitStyle("H1_czarny", "Nagłówek H1", StyleValues.Paragraph, "Nagłówek1", "48", "000000", bold: true),
        InitStyle("H1_pomaranczowy", "Nagłówek H1", StyleValues.Paragraph, "Nagłówek1", "48", "FF6200", bold: true),
        InitStyle("H1_zmienna", "Nagłówek H1", StyleValues.Paragraph, "Nagłówek1", "48", "0000FF", bold: true),

        InitStyle("H2_czarny", "Nagłówek H2", StyleValues.Paragraph, "Nagłówek2", "32", "000000", bold: true),
        InitStyle("H2_zmienna", "Nagłówek H2", StyleValues.Paragraph, "Nagłówek2", "32", "0000FF", bold: true),

        InitStyle("H3_czarny", "Nagłówek H3", StyleValues.Paragraph, "Nagłówek3", "28", "000000", bold: true),
        InitStyle("H3_zmienna", "Nagłówek H3", StyleValues.Paragraph, "Nagłówek3", "28", "0000FF", bold: true),

        InitStyle("H4_czarny", "Nagłówek H4", StyleValues.Paragraph, "Nagłówek4", "24", "000000", bold: true),
        InitStyle("H4_zmienna", "Nagłówek H4", StyleValues.Paragraph, "Nagłówek4", "24", "0000FF", bold: true),

        InitStyle("H5_czarny", "Nagłówek H5", StyleValues.Paragraph, "Nagłówek5", "20", "000000", bold: true),
        InitStyle("H5_zmienna", "Nagłówek H5", StyleValues.Paragraph, "Nagłówek5", "20", "0000FF", bold: true),

        // Spis treści
        InitStyle("TOCHeading", "TOC Heading", StyleValues.Paragraph, "Nagłówek1", "32", "000000"),

        // Paragraf
        InitStyle("Paragraf_regularny", "Paragraf", StyleValues.Paragraph, "Normalny", "20", "000000"),
        InitStyle("Paragraf_bold", "Paragraf", StyleValues.Paragraph, "Normalny", "20", "000000", bold: true),

        // Paragraf - zmienna
        InitStyle("Paragraf_zmienna", "Paragraf - zmienna", StyleValues.Paragraph, "Normalny", "20", "0000FF"),
        InitStyle("Paragraf_zmienna_bold", "Paragraf - zmienna", StyleValues.Paragraph, "Normalny", "20", "0000FF", bold: true),

        // Link
        InitStyle("Link", "Link", StyleValues.Paragraph, "HTML", "20", "0000FF", underline: true),

        // Komentarze do zapisów wariantowych
        InitStyle("KomentarzeZapisow", "Komentarze do zapisów wariantowych", StyleValues.Paragraph, "Normalny", "20", "FF0000", italic: true),

        // Zapisy zmienne i wariantowe
        InitStyle("ZapisyWariantowe", "Zapisy zmienne i wariantowe", StyleValues.Paragraph, "Normalny", "20", "696969"),

        // Tabela nagłówek i tabela wiersze
        InitStyle("TabelaNaglowek", "Tabela nagłówek", StyleValues.Paragraph, "Normalny", "20", "000000", bold: true),
        InitStyle("TabelaWiersze", "Tabela wiersze", StyleValues.Paragraph, "Normalny", "20", "000000"),

        // Podpis zmienna
        InitStyle("PodpisZmienna", "Podpis zmienna", StyleValues.Paragraph, "Normalny", "20", "0000FF"),

        // Stopka
        InitStyle("Stopka", "Stopka", StyleValues.Paragraph, "Normalny", "16", "000000", lineSpacing: "240"), // 1.2
        InitStyle("StopkaAdres", "Stopka - adres", StyleValues.Paragraph, "Normalny", "16", "000000", lineSpacing: "300") // 1.5
    };
}

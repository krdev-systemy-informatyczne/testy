
InitStyle("Heading2", "Nagłówek H2", StyleValues.Paragraph, "Heading2", "32", "0000FF", true),
InitStyle("Heading3", "Nagłówek H3", StyleValues.Paragraph, "Heading3", "28", "0000FF", true),
InitStyle("Heading4", "Nagłówek H4", StyleValues.Paragraph, "Heading4", "24", "0000FF", true),
InitStyle("Heading5", "Nagłówek H5", StyleValues.Paragraph, "Heading5", "20", "0000FF", true),

InitStyle("TOCHeading", "Nagłówek spisu treści", StyleValues.Paragraph, "TOCHeading", "20", "000000", true),

InitStyle("Paragraf", "Paragraf", StyleValues.Paragraph, "Normalny", "20", "000000", false),
InitStyle("ParagrafBold", "Paragraf (Bold)", StyleValues.Paragraph, "Normalny", "20", "000000", true),

InitStyle("ParagrafZmienna", "Paragraf - zmienna", StyleValues.Paragraph, "Normalny", "20", "0000FF", false),
InitStyle("ParagrafZmiennaBold", "Paragraf - zmienna (Bold)", StyleValues.Paragraph, "Normalny", "20", "0000FF", true),

InitStyle("Hyperlink", "Linki", StyleValues.Paragraph, "Normalny", "20", "0000FF", false, underline: true),

InitStyle("ZapisyZmienne", "Zapisy zmienne i wariantowe", StyleValues.Paragraph, "Normalny", "20", "0000FF", false),

InitStyle("KomentarzeZapisow", "Komentarze do zapisów wariantowych", StyleValues.Paragraph, "Normalny", "20", "C00060", false),

InitStyle("Footer", "Stopka", StyleValues.Paragraph, "Normalny", "16", "000000", false, lineSpacing: "240"),
InitStyle("FooterAddress", "Stopka - adres", StyleValues.Paragraph, "Normalny", "16", "000000", false, lineSpacing: "300"), // dla 1.5 linia

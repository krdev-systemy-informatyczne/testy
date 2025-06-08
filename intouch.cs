private WordprocessingDocument CreateDocumentWithShape(string altText)
{
    var wordDoc = WordprocessingDocument.Create(new MemoryStream(), WordprocessingDocumentType.Document);
    var mainPart = wordDoc.AddMainDocumentPart();
    mainPart.Document = new Document(new Body());

    // Tworzymy pic:pic
    var pic = new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties
            {
                Id = 1U,
                Name = "Shape1",
                Description = altText
            },
            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()
        ),
        new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(),
        new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties()
    );

    // Owijamy w a:graphicData
    var graphicData = new DocumentFormat.OpenXml.Drawing.GraphicData();
    graphicData.AppendChild(pic);

    // Owijamy w a:graphic
    var graphic = new DocumentFormat.OpenXml.Drawing.Graphic();
    graphic.AppendChild(graphicData);

    // Owijamy w wp:inline
    var inline = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline();
    inline.AppendChild(graphic);

    // Finalnie w DrawingML
    var drawing = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Drawing();
    drawing.AppendChild(inline);

    // Dodajemy do paragrafu
    mainPart.Document.Body.AppendChild(new Paragraph(new Run(drawing)));

    return wordDoc;
}

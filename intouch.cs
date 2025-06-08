using NUnit.Framework;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Pictures;
using D2DocumentQuality.WcagStandard;

[TestFixture]
public class AltTextWcagStandardRuleTests
{
    private AltTextWcagStandardRule _rule;

    [SetUp]
    public void Setup()
    {
        _rule = new AltTextWcagStandardRule();
    }

    [TestCase("", false, QualityIssueCode.MissingAltText, TestName = "ImageWithoutAltText_ShouldReportMissingAltText")]
    [TestCase("Valid alt text", false, null, TestName = "ImageWithAltText_ShouldPass")]
    [TestCase("", true, null, TestName = "DecorativeImage_ShouldPass")]
    public void CheckImagesAltText_VariousCases(string altText, bool isDecorative, QualityIssueCode? expectedIssueCode)
    {
        using (var wordDoc = CreateDocumentWithDrawing(altText, isDecorative))
        {
            var result = _rule.CheckImagesAltText(wordDoc.MainDocumentPart);
            AssertIssuesContainExpected(result, expectedIssueCode);
        }
    }

    [TestCase("", false, QualityIssueCode.DecorativeImageNoFlag, TestName = "DecorativeImageWithoutFlag_ShouldReport")]
    [TestCase("", true, null, TestName = "ProperDecorativeImage_ShouldPass")]
    public void CheckDecorativeImages_VariousCases(string altText, bool isDecorative, QualityIssueCode? expectedIssueCode)
    {
        using (var wordDoc = CreateDocumentWithDrawing(altText, isDecorative))
        {
            var result = _rule.CheckDecorativeImages(wordDoc.MainDocumentPart);
            AssertIssuesContainExpected(result, expectedIssueCode);
        }
    }

    [TestCase("Duplicate ALT", "Duplicate ALT", QualityIssueCode.AltTextDuplicate, TestName = "LinkedAltText_ShouldReportDuplicate")]
    [TestCase("ALT", "Different text", null, TestName = "LinkedAltText_ShouldPass")]
    public void CheckLinkedAltText_VariousCases(string altText, string neighborText, QualityIssueCode? expectedIssueCode)
    {
        using (var wordDoc = CreateDocumentWithDrawingAndNeighborText(altText, neighborText))
        {
            var result = _rule.CheckLinkedAltText(wordDoc.MainDocumentPart);
            AssertIssuesContainExpected(result, expectedIssueCode);
        }
    }

    [TestCase("", QualityIssueCode.HiddenAnnotation, TestName = "EmptyComment_ShouldReport")]
    [TestCase("Not empty", null, TestName = "FilledComment_ShouldPass")]
    public void CheckAnnotationHiding_VariousCases(string commentText, QualityIssueCode? expectedIssueCode)
    {
        using (var wordDoc = WordprocessingDocument.Create(new MemoryStream(), WordprocessingDocumentType.Document))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments = new Comments();
            commentsPart.Comments.AppendChild(new Comment() { InnerText = commentText });

            var result = _rule.CheckAnnotationHiding(mainPart);
            AssertIssuesContainExpected(result, expectedIssueCode);
        }
    }

    [TestCase("", QualityIssueCode.MissingAltText, TestName = "ShapeWithoutAltText_ShouldReport")]
    [TestCase("Valid shape", null, TestName = "ShapeWithAltText_ShouldPass")]
    public void CheckOtherAltTexts_VariousCases(string altText, QualityIssueCode? expectedIssueCode)
    {
        using (var wordDoc = CreateDocumentWithShape(altText))
        {
            var result = _rule.CheckOtherAltTexts(wordDoc.MainDocumentPart);
            AssertIssuesContainExpected(result, expectedIssueCode);
        }
    }

    private void AssertIssuesContainExpected(QualityResult result, QualityIssueCode? expectedIssueCode)
    {
        bool actualIssueFound = expectedIssueCode != null &&
                                result.Issues.Any(i => i.Code == expectedIssueCode);

        Assert.Multiple(() =>
        {
            if (expectedIssueCode != null)
            {
                Assert.IsTrue(actualIssueFound,
                    $"Expected issue code '{expectedIssueCode}' was not found.");
            }
            else
            {
                Assert.IsFalse(result.Issues.Any(),
                    $"Expected no issues, but found: [{string.Join(", ", result.Issues.Select(i => i.Code.ToString()))}]");
            }
        });
    }

    private WordprocessingDocument CreateDocumentWithDrawing(string altText, bool isDecorative = false)
    {
        var doc = WordprocessingDocument.Create(new MemoryStream(), WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var docPr = new DocProperties { Id = 1U, Name = "TestImage", Description = altText };

        if (isDecorative)
        {
            var decorativeAttr = new OpenXmlAttribute("decorative", null, "1");
            docPr.SetAttribute(decorativeAttr);
        }

        var drawing = new Drawing(new Inline(docPr));
        mainPart.Document.Body.AppendChild(new Paragraph(new Run(drawing)));

        return doc;
    }

    private WordprocessingDocument CreateDocumentWithDrawingAndNeighborText(string altText, string neighborText)
    {
        var doc = WordprocessingDocument.Create(new MemoryStream(), WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var docPr = new DocProperties { Id = 1U, Name = "TestImage", Description = altText };
        var drawing = new Drawing(new Inline(docPr));

        mainPart.Document.Body.Append(
            new Paragraph(new Run(drawing)),
            new Paragraph(new Run(new Text(neighborText)))
        );

        return doc;
    }

    private WordprocessingDocument CreateDocumentWithShape(string altText)
    {
        var doc = WordprocessingDocument.Create(new MemoryStream(), WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var pic = new Picture(
            new NonVisualPictureProperties(
                new NonVisualDrawingProperties
                {
                    Id = 1U,
                    Name = "Shape1",
                    Description = altText
                },
                new NonVisualPictureDrawingProperties()
            ),
            new BlipFill(),
            new ShapeProperties()
        );

        var graphicData = new GraphicData();
        graphicData.Append(pic);

        var graphic = new Graphic();
        graphic.Append(graphicData);

        var inline = new Inline();
        inline.Append(graphic);

        var drawingElement = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
        drawingElement.Append(inline);

        mainPart.Document.Body.AppendChild(new Paragraph(new Run(drawingElement)));

        return doc;
    }
}

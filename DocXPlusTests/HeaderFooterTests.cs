using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace DocXPlusTests
{
    [TestClass]
    public class HeaderFooterTests : TestBase
    {
        [TestMethod]
        public void AddFooter()
        {
            var filename = Path.Combine(TempDirectory, "AddFooter.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            var Footer = doc.AddFooter();
            Footer.AddParagraph().Append("Footer Paragraph");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddHeader()
        {
            var filename = Path.Combine(TempDirectory, "AddHeader.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            var header = doc.AddHeader();
            header.AddParagraph().Append("Header Paragraph");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddHeaderAndFooter()
        {
            var filename = Path.Combine(TempDirectory, "AddHeaderAndFooter.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            var header = doc.AddHeader();
            header.AddParagraph().Append("Header Paragraph");

            var footer = doc.AddFooter();
            footer.AddParagraph().Append("Footer Paragraph");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddHeaderAndFooterLandscape()
        {
            var filename = Path.Combine(TempDirectory, "AddHeaderAndFooterLandscape.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            doc.Orientation = DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues.Landscape;

            var header = doc.AddHeader();
            header.AddParagraph()
                .SetAlignment(DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Right)
                .Append(LoremIpsum);

            var footer = doc.AddFooter();
            footer.AddParagraph()
                .SetAlignment(DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center)
                .Append(LoremIpsum);

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }
    }
}
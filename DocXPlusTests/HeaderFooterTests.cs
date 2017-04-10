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

            var Footer = doc.AddFooter(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default);
            Footer.AddParagraph().Append("Footer Paragraph");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddFooters()
        {
            var filename = Path.Combine(TempDirectory, "AddFooters.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            doc.AddFooter(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default)
                .AddParagraph()
                .Append("Default (Odd) Footer");

            doc.AddFooter(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Even)
                .AddParagraph()
                .Append("Even Footer");

            doc.AddFooter(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.First)
                .AddParagraph()
                .Append("First Footer");

            doc.AddParagraph().Append("Page 1");

            doc.InsertPageBreak();

            doc.AddParagraph().Append("Page 2");

            doc.InsertPageBreak();

            doc.AddParagraph().Append("Page 3");

            doc.InsertPageBreak();

            doc.AddParagraph().Append("Page 4");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddHeader()
        {
            var filename = Path.Combine(TempDirectory, "AddHeader.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            var header = doc.AddHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default);
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

            var header = doc.AddHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default);
            header.AddParagraph().Append("Header Paragraph");

            var footer = doc.AddFooter(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default);
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

            var header = doc.AddHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default);

            header.AddParagraph()
                .SetAlignment(DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Right)
                .Append(LoremIpsum);

            var footer = doc.AddFooter(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default);

            footer.AddParagraph()
                .SetAlignment(DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center)
                .Append(LoremIpsum);

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddHeaders()
        {
            var filename = Path.Combine(TempDirectory, "AddHeaders.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            doc.AddHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default)
                .AddParagraph()
                .Append("Default (Odd) Header");

            doc.AddHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Even)
                .AddParagraph()
                .Append("Even Header");

            doc.AddHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.First)
                .AddParagraph()
                .Append("First Header");

            doc.AddParagraph().Append("Page 1");

            doc.InsertPageBreak();

            doc.AddParagraph().Append("Page 2");

            doc.InsertPageBreak();

            doc.AddParagraph().Append("Page 3");

            doc.InsertPageBreak();

            doc.AddParagraph().Append("Page 4");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddSectionFooter()
        {
            var filename = Path.Combine(TempDirectory, "AddSectionFooter.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            doc.AddFooter(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Footer 1");

            doc.InsertSectionPageBreak();

            doc.AddFooter(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Footer 2");

            doc.InsertSectionPageBreak();

            doc.AddFooter(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Footer 3");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddSectionHeader()
        {
            var filename = Path.Combine(TempDirectory, "AddSectionHeader.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            doc.AddHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Header 1");

            doc.InsertSectionPageBreak();

            doc.AddHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Header 2");

            doc.InsertSectionPageBreak();

            doc.AddHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Header 3");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddSectionHeaderFooter()
        {
            var filename = Path.Combine(TempDirectory, "AddSectionHeaderFooter.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            doc.AddHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Header 1");
            doc.AddFooter(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Footer 1");

            doc.InsertSectionPageBreak();

            doc.AddHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Header 2");
            doc.AddFooter(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Footer 2");

            doc.InsertSectionPageBreak();

            doc.AddHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Header 3");
            doc.AddFooter(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Footer 3");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddSectionPageBreakSameHeader()
        {
            var filename = Path.Combine(TempDirectory, "AddSectionPageBreakSameHeader.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            var header = doc.AddHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default);
            header.AddParagraph().Append("Header 1");

            doc.InsertSectionPageBreak();

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }
    }
}
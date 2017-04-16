using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocXPlus;
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

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

            doc.AddFooters();
            doc.DefaultFooter.AddParagraph().Append("Footer Paragraph");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddFooters()
        {
            var filename = Path.Combine(TempDirectory, "AddFooters.docx");

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

            doc.AddFooters();

            doc.DefaultFooter
                .AddParagraph()
                .Append("Default (Odd) Footer");

            doc.EvenFooter
                .AddParagraph()
                .Append("Even Footer");

            doc.FirstFooter
                .AddParagraph()
                .Append("First Footer");

            doc.DifferentFirstPage = true;
            doc.EvenAndOddHeaders = true;

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
        public void AddFooterWithNormalPageNumbers()
        {
            var filename = Path.Combine(TempDirectory, "AddFooterWithNormalPageNumbers.docx");

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

            doc.AddFooters();
            doc.DefaultFooter
                .AddParagraph()
                .Append("Page ")
                .AppendPageNumber(PageNumberFormat.Normal)
                .Append(" of ")
                .AppendPageCount(PageNumberFormat.Normal)
                .Bold()
                .Alignment = Align.Center;

            for (int i = 0; i < 9; i++)
            {
                doc.InsertPageBreak();
            }

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddFooterWithRomanPageNumbers()
        {
            var filename = Path.Combine(TempDirectory, "AddFooterWithRomanPageNumbers.docx");

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

            doc.AddFooters();
            doc.DefaultFooter
                .AddParagraph()
                .Append("Page: ")
                .AppendPageNumber(PageNumberFormat.Roman)
                .Append(" of ")
                .AppendPageCount(PageNumberFormat.Roman)
                .Bold()
                .Alignment = Align.Center;

            for (int i = 0; i < 9; i++)
            {
                doc.InsertPageBreak();
            }

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddHeader()
        {
            var filename = Path.Combine(TempDirectory, "AddHeader.docx");

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

            doc.AddHeaders();
            doc.DefaultHeader.AddParagraph().Append("Header Paragraph");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddHeaderAndFooter()
        {
            var filename = Path.Combine(TempDirectory, "AddHeaderAndFooter.docx");

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

            doc.AddHeaders();
            doc.AddFooters();

            doc.DefaultHeader.AddParagraph().Append("Header Paragraph");
            doc.DefaultFooter.AddParagraph().Append("Footer Paragraph");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddHeaderAndFooterLandscape()
        {
            var filename = Path.Combine(TempDirectory, "AddHeaderAndFooterLandscape.docx");

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);
            doc.Orientation = PageOrientationValues.Landscape;

            doc.AddHeaders();
            doc.AddFooters();

            doc.DefaultHeader
                .AddParagraph()
                .SetAlignment(Align.Right)
                .Append(LoremIpsum);

            doc.DefaultFooter
                .AddParagraph()
                .SetAlignment(Align.Center)
                .Append(LoremIpsum);

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddHeaders()
        {
            var filename = Path.Combine(TempDirectory, "AddHeaders.docx");

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

            doc.AddHeaders();
            doc.DifferentFirstPage = true;
            doc.EvenAndOddHeaders = true;

            doc.DefaultHeader
                .AddParagraph()
                .Append("Default (Odd) Header");

            doc.EvenHeader
                .AddParagraph()
                .Append("Even Header");

            doc.FirstHeader
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

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

            doc.AddFooters();

            doc.DefaultFooter.AddParagraph().Append("Footer 1");

            doc.InsertSectionPageBreak();

            doc.AddFooters();

            doc.DefaultFooter.AddParagraph().Append("Footer 2");

            doc.InsertSectionPageBreak();

            doc.AddFooters();

            doc.DefaultFooter.AddParagraph().Append("Footer 3");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddSectionHeader()
        {
            var filename = Path.Combine(TempDirectory, "AddSectionHeader.docx");

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

            doc.AddHeaders();

            doc.DefaultHeader.AddParagraph().Append("Header 1");

            doc.InsertSectionPageBreak();

            doc.AddHeaders();

            doc.DefaultHeader.AddParagraph().Append("Header 2");

            doc.InsertSectionPageBreak();

            doc.AddHeaders();

            doc.DefaultHeader.AddParagraph().Append("Header 3");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddSectionHeaderFooter()
        {
            var filename = Path.Combine(TempDirectory, "AddSectionHeaderFooter.docx");

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

            doc.AddHeaders();
            doc.AddFooters();

            doc.DefaultHeader.AddParagraph().Append("Header 1");
            doc.DefaultFooter.AddParagraph().Append("Footer 1");

            doc.InsertSectionPageBreak();

            doc.AddHeaders();
            doc.AddFooters();

            doc.DefaultHeader.AddParagraph().Append("Header 2");
            doc.DefaultFooter.AddParagraph().Append("Footer 2");

            doc.InsertSectionPageBreak();

            doc.AddHeaders();
            doc.AddFooters();

            doc.DefaultHeader.AddParagraph().Append("Header 3");
            doc.DefaultFooter.AddParagraph().Append("Footer 3");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddSectionPageBreakSameHeader()
        {
            var filename = Path.Combine(TempDirectory, "AddSectionPageBreakSameHeader.docx");

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

            doc.AddHeaders();

            doc.DefaultHeader.AddParagraph().Append("Header 1");

            doc.InsertSectionPageBreak();

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }
    }
}
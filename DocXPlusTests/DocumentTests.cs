using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace DocXPlusTests
{
    [TestClass]
    public class DocumentTests : TestBase
    {
        [TestMethod]
        public void CreateUsingFile()
        {
            var filename = Path.Combine(TempDirectory, "CreateUsingFile.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void Landscape()
        {
            var filename = Path.Combine(TempDirectory, "Landscape.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            doc.Orientation = DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues.Landscape;

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void SectionPageBreak()
        {
            var filename = Path.Combine(TempDirectory, "SectionPageBreak.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            doc.AddParagraph().Append("Page 1");

            doc.InsertSectionPageBreak();

            doc.AddParagraph().Append("Page 2");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }
    }
}
using DocXPlus;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace DocXPlusTests
{
    [TestClass]
    public class DocumentTests : TestBase
    {
        [TestMethod]
        public void CreateUsingDefaultStream()
        {
            var filename = Path.Combine(TempDirectory, "CreateUsingDefaultStream.docx");

            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddParagraph().Append(LoremIpsum);

                using (var stream = new FileStream(filename, FileMode.Create))
                {
                    doc.SaveAs(stream);
                }
            }

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void CreateUsingFile()
        {
            var filename = Path.Combine(TempDirectory, "CreateUsingFile.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            doc.AddParagraph().Append(LoremIpsum);

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void CreateUsingStream()
        {
            var filename = Path.Combine(TempDirectory, "CreateUsingStream.docx");

            using (var stream = new FileStream(filename, FileMode.Create))
            {
                var doc = DocX.Create(stream, DocumentType.Document);

                doc.AddParagraph().Append(LoremIpsum);

                doc.Close();
            }

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void Landscape()
        {
            var filename = Path.Combine(TempDirectory, "Landscape.docx");

            var doc = DocX.Create(filename, DocumentType.Document);
            doc.Orientation = PageOrientation.Landscape;

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }
    }
}
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace DocXPlusTests
{
    [TestClass]
    public class TableTests : TestBase
    {
        [TestMethod]
        public void ThreeColumnTable()
        {
            var filename = Path.Combine(TempDirectory, "ThreeColumnTable.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            var table = doc.AddTable(3);

            table.AddRow();
            table.AddRow();
            table.AddRow();

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }
    }
}
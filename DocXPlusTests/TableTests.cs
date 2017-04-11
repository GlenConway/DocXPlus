using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Linq;

namespace DocXPlusTests
{
    [TestClass]
    public class TableTests : TestBase
    {
        [TestMethod]
        public void ThreeByThreeTable()
        {
            var filename = Path.Combine(TempDirectory, "ThreeByThreeTable.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            var table = doc.AddTable(3);

            var row = table.AddRow();
            row.HeaderRow = true;

            row.Cells[0].Paragraphs[0].Append("Cell 1");
            row.Cells[1].Paragraphs[0].Append("Cell 2").SetAlignment(DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center);
            row.Cells[2].Paragraphs[0].Append("Cell 3").SetAlignment(DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Right);

            row = table.AddRow();

            row.Cells[0].Paragraphs[0].Append("Cell 1");
            row.Cells[1].Paragraphs[0].Append("Cell 2").SetAlignment(DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center);
            row.Cells[2].Paragraphs[0].Append("Cell 3").SetAlignment(DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Right);

            row = table.AddRow();

            row.Cells[0].Paragraphs[0].Append("Cell 1");
            row.Cells[1].Paragraphs[0].Append("Cell 2").SetAlignment(DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center);
            row.Cells[2].Paragraphs[0].Append("Cell 3").SetAlignment(DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Right);

            var rows = table.Rows.ToList();

            for (int i = 0; i < rows.Count(); i++)
            {
                row = rows[i];

                row.Height = DocXPlus.Units.UHalfInch;
                row.CantSplit = true;

                for (int j = 0; j < row.Cells.Count(); j++)
                {
                    var cell = row.Cells[j];

                    cell.SetVerticalAlignment(DocumentFormat.OpenXml.Wordprocessing.TableVerticalAlignmentValues.Center);
                }
            }

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }
    }
}
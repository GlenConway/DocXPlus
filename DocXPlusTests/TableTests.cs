using DocumentFormat.OpenXml.Wordprocessing;
using DocXPlus;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Linq;

namespace DocXPlusTests
{
    [TestClass]
    public class TableTests : TestBase
    {
        [TestMethod]
        public void TableWithHeaderRow()
        {
            var filename = Path.Combine(TempDirectory, "TableWithHeaderRow.docx");

            var doc = DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            var table = doc.AddTable(5);

            for (int i = 0; i < 50; i++)
            {
                var row = table.AddRow();
                row.SetBorders(Units.HalfPt, BorderValues.Single);

                if (i == 0)
                {
                    row.SetShading(ShadingPatternValues.Clear, "E7E6E6");

                    row.HeaderRow = true;
                }

                for (int j = 0; j < 5; j++)
                {
                    row.Cells[j].Paragraphs[0].Append($"Cell {(j + 1)}");
                }
            }

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }
        [TestMethod]
        public void TableWithMergeRight()
        {
            var filename = Path.Combine(TempDirectory, "TableWithMergeRight.docx");

            var doc = DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            var table = doc.AddTable(5);

            for (int i = 0; i < 5; i++)
            {
                var row = table.AddRow();
                row.SetBorders(Units.HalfPt, BorderValues.Single);

                if (i == 0)
                {
                    row.SetShading(ShadingPatternValues.Clear, "E7E6E6");

                    row.HeaderRow = true;
                }

                for (int j = 0; j < 5; j++)
                {
                    row.Cells[j].Paragraphs[0].Append($"Cell {(j + 1)}");
                }
            }

            var firstRow = table.Rows.First();
            firstRow.Cells[0].MergeRight = 3;
            firstRow.Cells[1].AddParagraph().Append("Should not display.");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }
        [TestMethod]
        public void ThreeByThreeTable()
        {
            var filename = Path.Combine(TempDirectory, "ThreeByThreeTable.docx");

            var doc = DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            var table = doc.AddTable(3);

            var row = table.AddRow();
            row.HeaderRow = true;

            row.Cells[0].Paragraphs[0].Append("Cell 1");
            row.Cells[1].Paragraphs[0].Append("Cell 2").SetAlignment(JustificationValues.Center);
            row.Cells[2].Paragraphs[0].Append("Cell 3").SetAlignment(JustificationValues.Right);

            row.Cells[0].Borders.Set(Units.HalfPt, BorderValues.Single);

            row.SetShading(ShadingPatternValues.Clear, "E7E6E6");

            row = table.AddRow();

            row.SetBorders(Units.HalfPt, BorderValues.Single);

            row.Cells[0].Paragraphs[0].Append("Cell 1");
            row.Cells[1].Paragraphs[0].Append("Cell 2").SetAlignment(JustificationValues.Center);
            row.Cells[2].Paragraphs[0].Append("Cell 3").SetAlignment(JustificationValues.Right);

            row = table.AddRow();

            row.Cells[0].Paragraphs[0].Append("Cell 1");
            row.Cells[1].Paragraphs[0].Append("Cell 2").SetAlignment(JustificationValues.Center);
            row.Cells[2].Paragraphs[0].Append("Cell 3").SetAlignment(JustificationValues.Right);

            row.Cells[1].Shading.Set(ShadingPatternValues.Clear, "F2F2F2");
            row.Cells[2].Shading.Set(ShadingPatternValues.Clear, "auto");

            var rows = table.Rows.ToList();

            for (int i = 0; i < rows.Count(); i++)
            {
                row = rows[i];

                row.Height = Units.UHalfInch;
                row.BreakAcrossPages = true;

                for (int j = 0; j < row.Cells.Count(); j++)
                {
                    var cell = row.Cells[j];

                    cell.SetVerticalAlignment(TableVerticalAlignmentValues.Center);
                }
            }

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }
    }
}
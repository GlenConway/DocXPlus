using DocumentFormat.OpenXml;
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
        public void TableInHeader()
        {
            var filename = Path.Combine(TempDirectory, "TableInHeader.docx");

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

            doc.AddHeaders();

            var table = doc.DefaultHeader.AddTable(3);

            for (int i = 0; i < 3; i++)
            {
                var row = table.AddRow();
                row.SetBorders(Units.HalfPt, BorderValues.Single);

                if (i == 0)
                {
                    row.SetShading(ShadingPatternValues.Clear, "E7E6E6");

                    row.HeaderRow = true;
                }

                for (int j = 0; j < 3; j++)
                {
                    row.Cells[j].Paragraphs[0].Append($"Cell {(j + 1)}");
                }
            }

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void TablePercent()
        {
            var filename = Path.Combine(TempDirectory, "TablePercent.docx");

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

            var table = doc.AddTable(3, 60, 20, 20);

            for (int i = 0; i < 3; i++)
            {
                var row = table.AddRow();
                row.SetBorders(Units.HalfPt, BorderValues.Single);

                if (i == 0)
                {
                    row.SetShading(ShadingPatternValues.Clear, "E7E6E6");

                    row.HeaderRow = true;
                }

                for (int j = 0; j < 3; j++)
                {
                    row.Cells[j].Paragraphs[0].Append($"Cell {(j + 1)}");
                }
            }

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void TableWidths()
        {
            var filename = Path.Combine(TempDirectory, "TableWidths.docx");

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

            var col1 = Units.InchToTwips(3.9).ToString();
            var col2 = Units.InchToTwips(1.3).ToString();

            var table = doc.AddTable(3, col1, col2, col2);

            for (int i = 0; i < 3; i++)
            {
                var row = table.AddRow();
                row.SetBorders(Units.HalfPt, BorderValues.Single);

                if (i == 0)
                {
                    row.SetShading(ShadingPatternValues.Clear, "E7E6E6");

                    row.HeaderRow = true;
                }

                for (int j = 0; j < 3; j++)
                {
                    row.Cells[j].Paragraphs[0].Append($"Cell {(j + 1)}");
                }
            }

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void TableWithHeaderRow()
        {
            var filename = Path.Combine(TempDirectory, "TableWithHeaderRow.docx");

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

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
        public void TableWithMergeDown()
        {
            var filename = Path.Combine(TempDirectory, "TableWithMergeDown.docx");

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

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

            var rows = table.Rows.ToList();

            var firstRow = rows[0];

            for (int i = 0; i < 5; i++)
            {
                if (i > 0)
                {
                    firstRow.Cells[i].MergeDown = i;
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

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

            // 5 columns
            var table = doc.AddTable(5);

            for (int i = 0; i < 5; i++)
            {
                var row = table.AddRow();

                row.SetBorders(Units.HalfPt, BorderValues.Single);

                if (i == 0)
                {
                    // shade the first row and set as a header
                    row.SetShading(ShadingPatternValues.Clear, "E7E6E6");

                    row.HeaderRow = true;
                }

                for (int j = 0; j < 5; j++)
                {
                    row.Cells[j].Paragraphs[0].Append($"Cell {(j + 1)}");
                }

                if (i > 0)
                {
                    row.Cells[0].MergeRight = i;
                }
            }

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void TableWithMergeRightAndDown()
        {
            var filename = Path.Combine(TempDirectory, "TableWithMergeRightAndDown.docx");

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

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

            var rows = table.Rows.ToList();

            var firstRow = rows[0];
            firstRow.Cells[0].MergeRight = 3;
            firstRow.Cells[1].AddParagraph().Append("Should not display.");

            var secondRow = rows[1];
            secondRow.Cells[1].MergeDown = 3;

            var thirdRow = rows[2];
            thirdRow.Cells[1].AddParagraph().Append("Should not display.");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void ThreeByThreeTable()
        {
            var filename = Path.Combine(TempDirectory, "ThreeByThreeTable.docx");

            var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

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
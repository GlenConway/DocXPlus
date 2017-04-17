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
        public void TableCellMargins()
        {
            var filename = Path.Combine(TempDirectory, "TableCellMargins.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            var table = doc.AddTable(5);

            var row = table.AddRow();

            row.Cells[0].Margins.LeftMargin.Width = Units.InchToTwips(.25).ToString();
            row.Cells[0].Margins.TopMargin.Width = Units.InchToTwips(.25).ToString();
            row.Cells[0].Margins.BottomMargin.Width = Units.InchToTwips(.25).ToString();
            row.Cells[0].Margins.RightMargin.Width = Units.InchToTwips(.25).ToString();

            row.Cells[0].SetText("1/4\"");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void TableInHeader()
        {
            var filename = Path.Combine(TempDirectory, "TableInHeader.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            doc.AddHeaders();

            var table = doc.DefaultHeader.AddTable(3);

            for (int i = 0; i < 3; i++)
            {
                var row = table.AddRow();
                row.SetBorders(Units.HalfPt, BorderValue.Single);

                if (i == 0)
                {
                    row.SetShading(ShadingPattern.Clear, "E7E6E6");

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
        public void TableInTable()
        {
            var filename = Path.Combine(TempDirectory, "TableInTable.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            var table = doc.AddTable(2, 50, 50);
            var row = table.AddRow();

            row.Cells[0].AddParagraph();
            row.Cells[1].AddParagraph();

            var table1 = row.Cells[0].AddTable(3);
            table1.AddRow().Cells[0].SetText("Table 1, Cell 1");

            var table2 = row.Cells[1].AddTable(3);
            table2.AddRow().Cells[0].SetText("Table 2, Cell 1");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void TableMargins()
        {
            var filename = Path.Combine(TempDirectory, "TableMargins.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            var table = doc.AddTable(5);

            table.DefaultMargins.LeftMargin.Width = System.Convert.ToInt16(Units.InchToTwips(.25).Value);
            table.DefaultMargins.TopMargin.Width = Units.InchToTwips(.25).ToString();
            table.DefaultMargins.BottomMargin.Width = Units.InchToTwips(.25).ToString();
            table.DefaultMargins.RightMargin.Width = System.Convert.ToInt16(Units.InchToTwips(.25).Value);

            var row = table.AddRow();

            row.Cells[0].SetText("1/4\"");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void TablePercent()
        {
            var filename = Path.Combine(TempDirectory, "TablePercent.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            var table = doc.AddTable(3, 60, 20, 20);

            for (int i = 0; i < 3; i++)
            {
                var row = table.AddRow();
                row.SetBorders(Units.HalfPt, BorderValue.Single);

                if (i == 0)
                {
                    row.SetShading(ShadingPattern.Clear, "E7E6E6");

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
        public void TableRows()
        {
            var filename = Path.Combine(TempDirectory, "TableRows.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            for (int k = 0; k < 5; k++)
            {
                var columnCount = (k + 1) * 2;

                var table = doc.AddTable(columnCount);

                for (int i = 0; i < columnCount; i++)
                {
                    var row = table.AddRow();
                    row.SetBorders(Units.HalfPt, BorderValue.Single);

                    if (i == 0)
                    {
                        row.SetShading(ShadingPattern.Clear, "E7E6E6");

                        row.HeaderRow = true;
                    }

                    for (int j = 0; j < columnCount; j++)
                    {
                        row.Cells[j].Paragraphs[0].Append($"Cell {(j + 1)}");
                    }
                }

                doc.AddParagraph();
            }

            var tables = doc.Tables.ToList();

            for (int i = 0; i < 5; i++)
            {
                var columnCount = (i + 1) * 2;

                var table = tables[i];

                Assert.IsNotNull(table);
                Assert.IsNotNull(table.Rows);
                Assert.AreEqual(columnCount, table.Rows.Count());
            }

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void TableWidthsCM()
        {
            var filename = Path.Combine(TempDirectory, "TableWidthsCM.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            var table = doc.AddTable(3, "9.906cm", "3.302cm", "3.302cm");

            for (int i = 0; i < 3; i++)
            {
                var row = table.AddRow();
                row.SetBorders(Units.HalfPt, BorderValue.Single);

                if (i == 0)
                {
                    row.SetShading(ShadingPattern.Clear, "E7E6E6");

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
        public void TableWidthsInches()
        {
            var filename = Path.Combine(TempDirectory, "TableWidthsInches.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            var table = doc.AddTable(3, "3.9in", "1.3in", "1.3in");

            for (int i = 0; i < 3; i++)
            {
                var row = table.AddRow();
                row.SetBorders(Units.HalfPt, BorderValue.Single);

                if (i == 0)
                {
                    row.SetShading(ShadingPattern.Clear, "E7E6E6");

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
        public void TableWidthsMixed()
        {
            var filename = Path.Combine(TempDirectory, "TableWidthsMixed.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            var col3 = Units.InchToTwips(1.3).ToString();

            var table = doc.AddTable(3, "9.906cm", "1.3in", col3);

            for (int i = 0; i < 3; i++)
            {
                var row = table.AddRow();
                row.SetBorders(Units.HalfPt, BorderValue.Single);

                if (i == 0)
                {
                    row.SetShading(ShadingPattern.Clear, "E7E6E6");

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
        public void TableWidthsTwips()
        {
            var filename = Path.Combine(TempDirectory, "TableWidthsTwips.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            var col1 = Units.InchToTwips(3.9).ToString();
            var col2 = Units.InchToTwips(1.3).ToString();

            var table = doc.AddTable(3, col1, col2, col2);

            for (int i = 0; i < 3; i++)
            {
                var row = table.AddRow();
                row.SetBorders(Units.HalfPt, BorderValue.Single);

                if (i == 0)
                {
                    row.SetShading(ShadingPattern.Clear, "E7E6E6");

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

            var doc = DocX.Create(filename, DocumentType.Document);

            var table = doc.AddTable(5);

            for (int i = 0; i < 50; i++)
            {
                var row = table.AddRow();
                row.SetBorders(Units.HalfPt, BorderValue.Single);

                if (i == 0)
                {
                    // shade the first row and set as a header
                    row.SetShading("E7E6E6");

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

            var doc = DocX.Create(filename, DocumentType.Document);

            var table = doc.AddTable(5);

            for (int i = 0; i < 5; i++)
            {
                var row = table.AddRow();
                row.SetBorders(Units.HalfPt, BorderValue.Single);

                if (i == 0)
                {
                    row.SetShading(ShadingPattern.Clear, "E7E6E6");

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

            var doc = DocX.Create(filename, DocumentType.Document);

            // 5 columns
            var table = doc.AddTable(5);

            for (int i = 0; i < 5; i++)
            {
                var row = table.AddRow();

                row.SetBorders(Units.HalfPt, BorderValue.Single);

                if (i == 0)
                {
                    // shade the first row and set as a header
                    row.SetShading(ShadingPattern.Clear, "E7E6E6");

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

            var doc = DocX.Create(filename, DocumentType.Document);

            var table = doc.AddTable(5);

            for (int i = 0; i < 5; i++)
            {
                var row = table.AddRow();
                row.SetBorders(Units.HalfPt, BorderValue.Single);

                if (i == 0)
                {
                    row.SetShading(ShadingPattern.Clear, "E7E6E6");

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
        public void TableWithMultipleMergeRight()
        {
            var filename = Path.Combine(TempDirectory, "TableWithMultipleMergeRight.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            var table = doc.AddTable(9);

            var row = table.AddRow();
            row.HeaderRow = true;

            for (int i = 0; i < 9; i++)
            {
                row.Cells[i].SetBoldText((i + 1).ToString());
            }

            row = table.AddRow();
            row.HeaderRow = true;

            // this should merge cells 2, 3 and 4
            row.Cells[1].MergeRight = 2;

            // this should merge cells 5 and 6
            row.Cells[4].MergeRight = 1;

            row.SetBoldText(1, "Test1", Align.Center);
            row.SetBoldText(4, "Test2", Align.Center);

            row.Underline(1, 4);

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void ThreeByThreeTable()
        {
            var filename = Path.Combine(TempDirectory, "ThreeByThreeTable.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            var table = doc.AddTable(3);

            var row = table.AddRow();
            row.HeaderRow = true;

            row.Cells[0].Paragraphs[0].Append("Cell 1");
            row.Cells[1].Paragraphs[0].Append("Cell 2").SetAlignment(Align.Center);
            row.Cells[2].Paragraphs[0].Append("Cell 3").SetAlignment(Align.Right);

            row.Cells[0].Borders.Set(Units.HalfPt, BorderValue.Single);

            row.SetShading(ShadingPattern.Clear, "E7E6E6");

            row = table.AddRow();

            row.SetBorders(Units.HalfPt, BorderValue.Single);

            row.Cells[0].Paragraphs[0].Append("Cell 1");
            row.Cells[1].Paragraphs[0].Append("Cell 2").SetAlignment(Align.Center);
            row.Cells[2].Paragraphs[0].Append("Cell 3").SetAlignment(Align.Right);

            row = table.AddRow();

            row.Cells[0].Paragraphs[0].Append("Cell 1");
            row.Cells[1].Paragraphs[0].Append("Cell 2").SetAlignment(Align.Center);
            row.Cells[2].Paragraphs[0].Append("Cell 3").SetAlignment(Align.Right);

            row.Cells[1].Shading.Set(ShadingPattern.Clear, "F2F2F2");
            row.Cells[2].Shading.Set(ShadingPattern.Clear, "auto");

            var rows = table.Rows.ToList();

            for (int i = 0; i < rows.Count(); i++)
            {
                row = rows[i];

                row.Height = Units.UHalfInch;
                row.BreakAcrossPages = true;

                for (int j = 0; j < row.Cells.Count(); j++)
                {
                    var cell = row.Cells[j];

                    cell.SetVerticalAlignment(TableVerticalAlignment.Center);
                }
            }

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void TwoColumnTableWithMergeRight()
        {
            var filename = Path.Combine(TempDirectory, "TwoColumnTableWithMergeRight.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            var table = doc.AddTable(2);
            var row = table.AddRow();
            row.Cells[0].MergeRight = 1;

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }
    }
}
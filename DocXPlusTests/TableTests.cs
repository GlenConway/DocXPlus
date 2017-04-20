using DocXPlus;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace DocXPlusTests
{
    [TestClass]
    public class TableTests : TestBase
    {
        [TestMethod]
        public void TableCellMargins()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                var table = doc.AddTable(5);

                var row = table.AddRow();

                row.Cells[0].Margins.LeftMargin.SetInch(.25);
                row.Cells[0].Margins.TopMargin.SetInch(.25);
                row.Cells[0].Margins.BottomMargin.SetInch(.25);
                row.Cells[0].Margins.RightMargin.SetInch(.25);

                row.Cells[0].SetText("1/4\"");

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void TableInHeader()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void TableInTable()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                var table = doc.AddTable(2, 50, 50);
                var row = table.AddRow();

                row.Cells[0].AddParagraph();
                row.Cells[1].AddParagraph();

                var table1 = row.Cells[0].AddTable(3);
                table1.AddRow().Cells[0].SetText("Table 1, Cell 1");

                var table2 = row.Cells[1].AddTable(3);
                table2.AddRow().Cells[0].SetText("Table 2, Cell 1");

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void TableMargins()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                var table = doc.AddTable(5);

                table.DefaultMargins.LeftMargin.Width = System.Convert.ToInt16(Units.InchToTwips(.25).Value);
                table.DefaultMargins.TopMargin.Width = Units.InchToTwips(.25).ToString();
                table.DefaultMargins.BottomMargin.Width = Units.InchToTwips(.25).ToString();
                table.DefaultMargins.RightMargin.Width = System.Convert.ToInt16(Units.InchToTwips(.25).Value);

                var row = table.AddRow();

                row.Cells[0].SetText("1/4\"");

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void TablePercent()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void TableRows()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void TableWidthsCM()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void TableWidthsInches()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void TableWidthsMixed()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void TableWidthsTwips()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void TableWithHeaderRow()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void TableWithMergeDown()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void TableWithMergeRight()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void TableWithMergeRightAndDown()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void TableWithMultipleMergeRight()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void ThreeByThreeTable()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void TwoColumnTableWithMergeRight()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                var table = doc.AddTable(2);
                var row = table.AddRow();
                row.Cells[0].MergeRight = 1;

                Validate(doc);

                doc.Close();
            }
        }
    }
}
using DocXPlus;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DocXPlusTests
{
    [TestClass]
    public class HeaderFooterTests : TestBase
    {
        [TestMethod]
        public void AddFooter()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddFooters();
                doc.DefaultFooter.AddParagraph().Append("Footer Paragraph");

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void AddFooters()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void AddFooterWithNormalPageNumbers()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddFooters();
                doc.DefaultFooter
                    .AddParagraph()
                    .Append("Page ")
                    .AppendPageNumber(PageNumberFormat.Normal)
                    .Append(" of ")
                    .AppendPageCount(PageNumberFormat.Normal)
                    .Alignment = Align.Center;

                for (int i = 0; i < 9; i++)
                {
                    doc.InsertPageBreak();
                }

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void AddFooterWithNormalBoldPageNumbers()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void AddFooterWithRomanPageNumbers()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void AddHeader()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddHeaders();
                doc.DefaultHeader.AddParagraph().Append("Header Paragraph");

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void AddHeaderAndFooter()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddHeaders();
                doc.AddFooters();

                doc.DefaultHeader.AddParagraph().Append("Header Paragraph");
                doc.DefaultFooter.AddParagraph().Append("Footer Paragraph");

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void AddHeaderAndFooterLandscape()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.Orientation = PageOrientation.Landscape;

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void AddHeaders()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void AddSectionFooter()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddFooters();

                doc.DefaultFooter.AddParagraph().Append("Footer 1");

                doc.InsertSectionPageBreak();

                doc.AddFooters();

                doc.DefaultFooter.AddParagraph().Append("Footer 2");

                doc.InsertSectionPageBreak();

                doc.AddFooters();

                doc.DefaultFooter.AddParagraph().Append("Footer 3");

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void AddSectionHeader()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddHeaders();

                doc.DefaultHeader.AddParagraph().Append("Header 1");

                doc.InsertSectionPageBreak();

                doc.AddHeaders();

                doc.DefaultHeader.AddParagraph().Append("Header 2");

                doc.InsertSectionPageBreak();

                doc.AddHeaders();

                doc.DefaultHeader.AddParagraph().Append("Header 3");

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void AddSectionHeaderFooter()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void AddSectionPageBreakSameHeader()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddHeaders();

                doc.DefaultHeader.AddParagraph().Append("Header 1");

                doc.InsertSectionPageBreak();

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void AddHeaderPortraitThenLandscape()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddHeaders();

                var table = doc.DefaultHeader.AddTable(3);
                table.WidthType = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Auto;

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

                doc.InsertSectionPageBreak();
                doc.Orientation = PageOrientation.Landscape;

                // default is to link previous header
                // but if the orientation changes, the table will
                // only be the width of the portrait page

                // in order to have a new width you have to unlink
                // the header to the previous
                doc.AddHeaders();

                // and recreate the header

                table = doc.DefaultHeader.AddTable(3);
                table.WidthType = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Auto;

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

                //doc.SaveAs(System.IO.Path.Combine(TempDirectory, "AddHeaderPortraitThenLandscape.docx"));

                doc.Close();
            }
        }
    }
}
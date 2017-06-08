using DocXPlus;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DocXPlusTests
{
    [TestClass]
    public class StyleTests : TestBase
    {
        [TestMethod]
        public void HeaderAndFooterStyle()
        {
            using (var doc = new DocX())
            {
                doc.Create();
                doc.Styles.DocumentStyle("Normal").Size = 8;

                doc.AddHeaders();
                doc.AddFooters();

                doc.DefaultHeader.AddParagraph().Append("Header Paragraph");
                doc.DefaultFooter.AddParagraph().Append("Footer Paragraph");

                Validate(doc);

                //doc.SaveAs(System.IO.Path.Combine(TempDirectory, "HeaderAndFooterStyle.docx"));

                doc.Close();
            }
        }

        [TestMethod]
        public void NormalStyle()
        {
            using (var doc = new DocX())
            {
                doc.Create();
                doc.Styles.Normal.Size = 8;

                doc.AddParagraph().Append("Normal Style");

                Validate(doc);

                //doc.SaveAs(System.IO.Path.Combine(TempDirectory, "NormalStyle.docx"));

                doc.Close();
            }
        }

        [TestMethod]
        public void NormalStyleByName()
        {
            using (var doc = new DocX())
            {
                doc.Create();
                doc.Styles.DocumentStyle("Normal").Size = 8;

                doc.AddParagraph().Append("Normal Style");

                Validate(doc);

                //doc.SaveAs(System.IO.Path.Combine(TempDirectory, "NormalStyleByName.docx"));

                doc.Close();
            }
        }

        [TestMethod]
        public void TableStyle()
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

                Validate(doc);

                doc.SaveAs(System.IO.Path.Combine(TempDirectory, "TableStyle.docx"));

                doc.Close();
            }
        }
    }
}
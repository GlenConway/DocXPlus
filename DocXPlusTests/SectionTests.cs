using DocXPlus;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DocXPlusTests
{
    [TestClass]
    public class SectionTests : TestBase
    {
        [TestMethod]
        public void SectionPageBreak()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddParagraph().Append("Page 1");

                doc.InsertSectionPageBreak();

                doc.AddParagraph().Append("Page 2");

                doc.InsertSectionPageBreak();

                doc.AddParagraph().Append("Page 3");

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void SectionPageBreakLandscapePortrait()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.Orientation = PageOrientation.Landscape;

                doc.AddParagraph().Append("Landscape");

                doc.InsertSectionPageBreak();

                doc.Orientation = PageOrientation.Portrait;

                doc.AddParagraph().Append("Portrait");

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void SectionPageBreakLandscapePortraitLandscape()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.Orientation = PageOrientation.Landscape;

                doc.AddParagraph().Append("Landscape");

                doc.InsertSectionPageBreak();
                doc.Orientation = PageOrientation.Portrait;

                doc.AddParagraph().Append("Portrait");

                doc.InsertSectionPageBreak();
                doc.Orientation = PageOrientation.Landscape;

                doc.AddParagraph().Append("Landscape");

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void SectionPageBreakPortraitLandscape()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddParagraph().Append("Portrait");

                doc.InsertSectionPageBreak();

                doc.Orientation = PageOrientation.Landscape;

                doc.AddParagraph().Append("Landscape");

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void SectionPageBreakPortraitLandscapePortrait()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddParagraph().Append("Portrait");

                doc.InsertSectionPageBreak();

                doc.Orientation = PageOrientation.Landscape;

                doc.AddParagraph().Append("Landscape");

                doc.InsertSectionPageBreak();

                doc.Orientation = PageOrientation.Portrait;

                doc.AddParagraph().Append("Portrait");

                Validate(doc);

                doc.Close();
            }
        }
    }
}
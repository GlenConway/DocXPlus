using DocXPlus;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DocXPlusTests
{
    [TestClass]
    public class ParagraphTests : TestBase
    {
        [TestMethod]
        public void BoldParagraphs()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddParagraph().Append("Append normal paragraph");

                doc.AddParagraph().Append("Append then set bold").Bold();

                doc.AddParagraph().AppendBold("Append bold paragraph");

                var paragraph = doc.AddParagraph();
                paragraph.Bold();
                paragraph.Append("Add paragraph, set bold then append text.");

                doc.AddParagraph().Append("Append normal paragraph").AppendBold("Then append bold paragraph");

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void Fonts()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddParagraph().Append("Append normal paragraph");

                doc.AddParagraph().Append("Append Arial").FontFamily("Arial");

                doc.AddParagraph().Append("Append 20 points").FontSize(40);

                doc.AddParagraph().Append("Append 20 points").StyleName = "Heading1";

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void ItalicParagraphs()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddParagraph().Append("Append normal paragraph");

                doc.AddParagraph().Append("Append then set Italic").Italic();

                doc.AddParagraph().AppendItalic("Append Italic paragraph");

                var paragraph = doc.AddParagraph();
                paragraph.Italic();
                paragraph.Append("Add paragraph, set Italic then append text.");

                doc.AddParagraph().Append("Append normal paragraph").AppendItalic("Then append Italic paragraph");

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void ParagraphAlignment()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddParagraph().AppendBold("Default (Left)");

                doc.AddParagraph().Append(LoremIpsum);

                doc.AddParagraph().AppendBold("Right");

                doc.AddParagraph().Append(LoremIpsum).SetAlignment(Align.Right);

                doc.AddParagraph().AppendBold("Center");

                var paragraph = doc.AddParagraph();
                paragraph.SetAlignment(Align.Center);
                paragraph.Append(LoremIpsum);

                doc.AddParagraph().AppendBold("Both");

                doc.AddParagraph().Append(LoremIpsum).SetAlignment(Align.Both);

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void ParagraphIndentation()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddParagraph().AppendBold("1\" IndentationBefore");

                doc.AddParagraph().Append(LoremIpsum).IndentationBefore = Units.InchToTwips(1);

                doc.AddParagraph().AppendBold("1\" IndentationFirstLine");

                doc.AddParagraph().Append(LoremIpsum).IndentationFirstLine = Units.InchToTwips(1);

                doc.AddParagraph().AppendBold("1\" IndentationHanging");

                doc.AddParagraph().Append(LoremIpsum).IndentationHanging = Units.InchToTwips(1);

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void Paragraphs()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddParagraph().Append("Append paragraph");

                doc.AddParagraph().Append("Append paragraph").Append("Append again");

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void UnderlineParagraphs()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddParagraph().Append("Append normal paragraph");

                doc.AddParagraph().Append("Append then set Underline").Underline(UnderlineType.Single);

                doc.AddParagraph().AppendUnderline("Append Underline paragraph", UnderlineType.Single);

                var paragraph = doc.AddParagraph();
                paragraph.Underline(UnderlineType.Single);
                paragraph.Append("Add paragraph, set Underline then append text.");

                doc.AddParagraph().Append("Append normal paragraph").AppendUnderline("Then append Underline paragraph", UnderlineType.Single);

                doc.AddParagraph().AppendUnderline("Dash", UnderlineType.Dash);
                doc.AddParagraph().AppendUnderline("DashDotDotHeavy", UnderlineType.DashDotDotHeavy);
                doc.AddParagraph().AppendUnderline("DashDotHeavy", UnderlineType.DashDotHeavy);
                doc.AddParagraph().AppendUnderline("DashedHeavy", UnderlineType.DashedHeavy);
                doc.AddParagraph().AppendUnderline("DashLong", UnderlineType.DashLong);
                doc.AddParagraph().AppendUnderline("DashLongHeavy", UnderlineType.DashLongHeavy);
                doc.AddParagraph().AppendUnderline("DotDash", UnderlineType.DotDash);
                doc.AddParagraph().AppendUnderline("DotDotDash", UnderlineType.DotDotDash);
                doc.AddParagraph().AppendUnderline("Dotted", UnderlineType.Dotted);
                doc.AddParagraph().AppendUnderline("DottedHeavy", UnderlineType.DottedHeavy);
                doc.AddParagraph().AppendUnderline("Double", UnderlineType.Double);
                doc.AddParagraph().AppendUnderline("None", UnderlineType.None);
                doc.AddParagraph().AppendUnderline("Single", UnderlineType.Single);
                doc.AddParagraph().AppendUnderline("Thick", UnderlineType.Thick);
                doc.AddParagraph().AppendUnderline("Wave", UnderlineType.Wave);
                doc.AddParagraph().AppendUnderline("WavyDouble", UnderlineType.WavyDouble);
                doc.AddParagraph().AppendUnderline("WavyHeavy", UnderlineType.WavyHeavy);
                doc.AddParagraph().AppendUnderline("Words Words Words", UnderlineType.Words);

                Validate(doc);

                doc.Close();
            }
        }
    }
}
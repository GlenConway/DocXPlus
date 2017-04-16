using DocXPlus;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace DocXPlusTests
{
    [TestClass]
    public class ParagraphTests : TestBase
    {
        [TestMethod]
        public void BoldParagraphs()
        {
            var filename = Path.Combine(TempDirectory, "BoldParagraphs.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            doc.AddParagraph().Append("Append normal paragraph");

            doc.AddParagraph().Append("Append then set bold").Bold();

            doc.AddParagraph().AppendBold("Append bold paragraph");

            var paragraph = doc.AddParagraph();
            paragraph.Bold();
            paragraph.Append("Add paragraph, set bold then append text.");

            doc.AddParagraph().Append("Append normal paragraph").AppendBold("Then append bold paragraph");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void Fonts()
        {
            var filename = Path.Combine(TempDirectory, "Fonts.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            doc.AddParagraph().Append("Append normal paragraph");

            doc.AddParagraph().Append("Append Arial").FontFamily("Arial");

            doc.AddParagraph().Append("Append 20 points").FontSize(40);

            doc.AddParagraph().Append("Append 20 points").StyleName = "Heading1";

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void ItalicParagraphs()
        {
            var filename = Path.Combine(TempDirectory, "ItalicParagraphs.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            doc.AddParagraph().Append("Append normal paragraph");

            doc.AddParagraph().Append("Append then set Italic").Italic();

            doc.AddParagraph().AppendItalic("Append Italic paragraph");

            var paragraph = doc.AddParagraph();
            paragraph.Italic();
            paragraph.Append("Add paragraph, set Italic then append text.");

            doc.AddParagraph().Append("Append normal paragraph").AppendItalic("Then append Italic paragraph");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void ParagraphAlignment()
        {
            var filename = Path.Combine(TempDirectory, "ParagraphAlignment.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

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

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void ParagraphIndentation()
        {
            var filename = Path.Combine(TempDirectory, "ParagraphIndentation.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            doc.AddParagraph().AppendBold("1\" IndentationBefore");

            doc.AddParagraph().Append(LoremIpsum).IndentationBefore = Units.InchToTwips(1);

            doc.AddParagraph().AppendBold("1\" IndentationFirstLine");

            doc.AddParagraph().Append(LoremIpsum).IndentationFirstLine = Units.InchToTwips(1);

            doc.AddParagraph().AppendBold("1\" IndentationHanging");

            doc.AddParagraph().Append(LoremIpsum).IndentationHanging = Units.InchToTwips(1);

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void Paragraphs()
        {
            var filename = Path.Combine(TempDirectory, "Paragraphs.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            doc.AddParagraph().Append("Append paragraph");

            doc.AddParagraph().Append("Append paragraph").Append("Append again");

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void UnderlineParagraphs()
        {
            var filename = Path.Combine(TempDirectory, "UnderlineParagraphs.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

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

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }
    }
}
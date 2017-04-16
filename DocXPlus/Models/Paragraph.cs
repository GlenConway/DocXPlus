using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DocXPlus
{
    /// <summary>
    /// Represents a Word paragraph
    /// </summary>
    public class Paragraph
    {
        private DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph;

        internal Paragraph(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph)
        {
            this.paragraph = paragraph;
        }

        /// <summary>
        /// Gets or sets the paragrapg alignment
        /// </summary>
        public Align Alignment
        {
            get
            {
                var justification = GetParagraphProperties().GetOrCreate<Justification>();
                return Utils.ConvertToAlign(justification.Val);
            }
            set
            {
                var justification = GetParagraphProperties().GetOrCreate<Justification>();
                justification.Val = Utils.ConvertToJustificationValues(value);
            }
        }

        /// <summary>
        /// Gets or sets the paragraph indentation in Twips
        /// </summary>
        public Int32Value IndentationBefore
        {
            get
            {
                if (!GetParagraphProperties().Has<Indentation>())
                    return 0;

                var indentation = GetParagraphProperties().GetOrCreate<Indentation>();

                if (int.TryParse(indentation.Left, out int result))
                    return result;

                return 0;
            }
            set
            {
                var indentation = GetParagraphProperties().GetOrCreate<Indentation>();

                indentation.Left = value.Value.ToString();
            }
        }

        /// <summary>
        /// Gets or sets the paragraph first line indentation in Twips
        /// </summary>
        public Int32Value IndentationFirstLine
        {
            get
            {
                if (!GetParagraphProperties().Has<Indentation>())
                    return 0;

                var indentation = GetParagraphProperties().GetOrCreate<Indentation>();

                if (int.TryParse(indentation.FirstLine, out int result))
                    return result;

                return 0;
            }
            set
            {
                var indentation = GetParagraphProperties().GetOrCreate<Indentation>();

                indentation.FirstLine = value.Value.ToString();
            }
        }

        /// <summary>
        /// Gets or sets the paragraph hanging indentation in Twips
        /// </summary>
        public Int32Value IndentationHanging
        {
            get
            {
                if (!GetParagraphProperties().Has<Indentation>())
                    return 0;

                var indentation = GetParagraphProperties().GetOrCreate<Indentation>();

                if (int.TryParse(indentation.Hanging, out int result))
                    return result;

                return 0;
            }
            set
            {
                var indentation = GetParagraphProperties().GetOrCreate<Indentation>();

                indentation.Hanging = value.Value.ToString();
            }
        }

        /// <summary>
        /// Gets or sets the name of the style associated with the paragraph
        /// </summary>
        public string StyleName
        {
            get
            {
                var style = GetParagraphProperties().GetOrCreate<ParagraphStyleId>();
                return style.Val;
            }
            set
            {
                var style = GetParagraphProperties().GetOrCreate<ParagraphStyleId>();
                style.Val = value;
            }
        }

        private IEnumerable<Run> Runs
        {
            get
            {
                return paragraph.Elements<Run>();
            }
        }

        /// <summary>
        /// Appends a Drawing object to the paragraph
        /// </summary>
        /// <param name="drawing"></param>
        /// <returns></returns>
        public Paragraph Append(Drawing drawing)
        {
            paragraph.AppendChild(new Run(drawing));

            return this;
        }

        /// <summary>
        /// Appends text to the paragraph
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public Paragraph Append(string text)
        {
            GetRun(text);

            return this;
        }

        /// <summary>
        /// Appends bold text to the paragraph
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public Paragraph AppendBold(string text)
        {
            var run = GetRun(text);
            run.Bold();

            return this;
        }

        /// <summary>
        /// Appends italic text to the paragraph
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public Paragraph AppendItalic(string text)
        {
            var run = GetRun(text);
            run.Italic();

            return this;
        }

        /// <summary>
        /// Appends a page count code to the paragraph
        /// </summary>
        /// <param name="format"></param>
        /// <returns></returns>
        public Paragraph AppendPageCount(PageNumberFormat format)
        {
            var run = paragraph.AppendChild(new Run());
            var fieldChar = run.GetOrCreate<FieldChar>();
            fieldChar.FieldCharType = FieldCharValues.Begin;

            run = paragraph.AppendChild(new Run());
            var fieldCode = run.GetOrCreate<FieldCode>();
            fieldCode.Space = SpaceProcessingModeValues.Preserve;

            if (format == PageNumberFormat.Normal)
            {
                fieldCode.Text = @" NUMPAGES   \* MERGEFORMAT ";
            }
            else
            {
                fieldCode.Text = @" NUMPAGES  \* ROMAN  \* MERGEFORMAT ";
            }

            run = paragraph.AppendChild(new Run());
            fieldChar = run.GetOrCreate<FieldChar>();
            fieldChar.FieldCharType = FieldCharValues.Separate;

            run = paragraph.AppendChild(new Run());
            var runProperties = run.GetOrCreate<RunProperties>();
            var noProof = runProperties.GetOrCreate<NoProof>();
            run.AppendChild(new Text("1"));

            run = paragraph.AppendChild(new Run());
            runProperties = run.GetOrCreate<RunProperties>();
            noProof = runProperties.GetOrCreate<NoProof>();
            fieldChar = run.GetOrCreate<FieldChar>();
            fieldChar.FieldCharType = FieldCharValues.End;

            return this;
        }

        /// <summary>
        /// Appends a page number code to the paragraph
        /// </summary>
        /// <param name="format"></param>
        /// <returns></returns>
        public Paragraph AppendPageNumber(PageNumberFormat format)
        {
            var run = paragraph.AppendChild(new Run());
            var fieldChar = run.GetOrCreate<FieldChar>();
            fieldChar.FieldCharType = FieldCharValues.Begin;

            run = paragraph.AppendChild(new Run());
            var fieldCode = run.GetOrCreate<FieldCode>();
            fieldCode.Space = SpaceProcessingModeValues.Preserve;

            if (format == PageNumberFormat.Normal)
            {
                fieldCode.Text = @" PAGE   \* MERGEFORMAT ";
            }
            else
            {
                fieldCode.Text = @" PAGE  \* ROMAN  \* MERGEFORMAT ";
            }

            run = paragraph.AppendChild(new Run());
            fieldChar = run.GetOrCreate<FieldChar>();
            fieldChar.FieldCharType = FieldCharValues.Separate;

            run = paragraph.AppendChild(new Run());
            var runProperties = run.GetOrCreate<RunProperties>();
            var noProof = runProperties.GetOrCreate<NoProof>();
            run.AppendChild(new Text("1"));

            run = paragraph.AppendChild(new Run());
            runProperties = run.GetOrCreate<RunProperties>();
            noProof = runProperties.GetOrCreate<NoProof>();
            fieldChar = run.GetOrCreate<FieldChar>();
            fieldChar.FieldCharType = FieldCharValues.End;

            return this;
        }

        /// <summary>
        /// Appends underlined text to the paragraph
        /// </summary>
        /// <param name="text"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public Paragraph AppendUnderline(string text, UnderlineValues value)
        {
            var run = GetRun(text);
            run.Underline(value);

            return this;
        }

        /// <summary>
        /// Makes a paragraph text bold
        /// </summary>
        /// <returns></returns>
        public Paragraph Bold()
        {
            if (Runs.Count() == 0)
            {
                var paragraphProperties = paragraph.GetOrCreate<ParagraphProperties>();
                var paragraphMarkRunProperties = paragraphProperties.GetOrCreate<ParagraphMarkRunProperties>();

                Bold bold = paragraphMarkRunProperties.GetOrCreate<Bold>();
                bold.Val = true;
            }
            else
            {
                foreach (var run in Runs)
                {
                    run.Bold();
                }
            }

            return this;
        }

        /// <summary>
        /// Applies the supplied font family to the paragraph
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public Paragraph FontFamily(string name)
        {
            if (Runs.Count() == 0)
            {
                var paragraphProperties = paragraph.GetOrCreate<ParagraphProperties>();
                var paragraphMarkRunProperties = paragraphProperties.GetOrCreate<ParagraphMarkRunProperties>();

                RunFonts prop = paragraphMarkRunProperties.GetOrCreate<RunFonts>();
                prop.Ascii = name;
                prop.HighAnsi = name;
                prop.ComplexScript = name;
                prop.EastAsia = name;
            }
            else
            {
                foreach (var run in Runs)
                {
                    run.FontFamily(name);
                }
            }

            return this;
        }

        /// <summary>
        /// Applies the supplied font size to the paragraph
        /// </summary>
        /// <param name="size">Size is in half points e.g. 40 is 20pt</param>
        /// <returns></returns>
        public Paragraph FontSize(double size)
        {
            double temp = size * 2;

            if (temp - (int)temp == 0)
            {
                if (!(size > 0 && size < 1639))
                {
                    throw new ArgumentOutOfRangeException(nameof(size), "Value must be in the range 0 - 1638");
                }
            }
            else
            {
                throw new ArgumentException(nameof(size), "Value must be either a whole or half number, examples: 32, 32.5");
            }

            if (Runs.Count() == 0)
            {
                var paragraphProperties = paragraph.GetOrCreate<ParagraphProperties>();
                var paragraphMarkRunProperties = paragraphProperties.GetOrCreate<ParagraphMarkRunProperties>();

                FontSize fontSize = paragraphMarkRunProperties.GetOrCreate<FontSize>();
                fontSize.Val = size.ToString();

                FontSizeComplexScript fontSizeComplexScript = paragraphMarkRunProperties.GetOrCreate<FontSizeComplexScript>();
                fontSizeComplexScript.Val = size.ToString();
            }
            else
            {
                foreach (var run in Runs)
                {
                    run.FontSize(size);
                }
            }

            return this;
        }

        /// <summary>
        /// Makes a paragraph text italic
        /// </summary>
        /// <returns></returns>
        public Paragraph Italic()
        {
            if (Runs.Count() == 0)
            {
                var paragraphProperties = paragraph.GetOrCreate<ParagraphProperties>();
                var paragraphMarkRunProperties = paragraphProperties.GetOrCreate<ParagraphMarkRunProperties>();

                Italic Italic = paragraphMarkRunProperties.GetOrCreate<Italic>();
                Italic.Val = OnOffValue.FromBoolean(true);
            }
            else
            {
                foreach (var run in Runs)
                {
                    run.Italic();
                }
            }

            return this;
        }

        /// <summary>
        /// Sets the paragraph alignment
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public Paragraph SetAlignment(Align value)
        {
            var justification = GetParagraphProperties().GetOrCreate<Justification>();
            justification.Val = Utils.ConvertToJustificationValues(value);

            return this;
        }

        /// <summary>
        /// Makes a paragraph text underlined
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public Paragraph Underline(UnderlineValues value)
        {
            if (Runs.Count() == 0)
            {
                var paragraphProperties = paragraph.GetOrCreate<ParagraphProperties>();
                var paragraphMarkRunProperties = paragraphProperties.GetOrCreate<ParagraphMarkRunProperties>();

                Underline Underline = paragraphMarkRunProperties.GetOrCreate<Underline>();
                Underline.Val = value;
            }
            else
            {
                foreach (var run in Runs)
                {
                    run.Underline(value);
                }
            }

            return this;
        }

        internal ParagraphProperties GetParagraphProperties()
        {
            return paragraph.GetOrCreate<ParagraphProperties>(true);
        }

        internal Run NewRun()
        {
            var result = new Run();

            if (paragraph.Has<ParagraphProperties>())
            {
                var paragraphProperties = paragraph.GetOrCreate<ParagraphProperties>();

                if (paragraphProperties.Has<ParagraphMarkRunProperties>())
                {
                    var paragraphMarkRunProperties = paragraphProperties.GetOrCreate<ParagraphMarkRunProperties>();

                    result.PrependChild(paragraphMarkRunProperties.CloneNode(true));
                }
            }

            return result;
        }

        private Run GetRun(string text)
        {
            var run = paragraph.AppendChild(NewRun());

            var t = new Text(text)
            {
                Space = SpaceProcessingModeValues.Preserve
            };

            run.AppendChild(t);

            return run;
        }
    }
}
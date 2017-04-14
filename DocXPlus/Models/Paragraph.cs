using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

namespace DocXPlus
{
    public class Paragraph
    {
        private DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph;

        public Paragraph(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph)
        {
            this.paragraph = paragraph;
        }

        private IEnumerable<Run> Runs
        {
            get
            {
                return paragraph.Elements<Run>();
            }
        }

        public Paragraph AddPageNumber(PageNumberFormat format)
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

        public Paragraph Append(Drawing drawing)
        {
            paragraph.AppendChild(new Run(drawing));

            return this;
        }

        public Paragraph Append(string text)
        {
            GetRun(text);

            return this;
        }

        public Paragraph AppendBold(string text)
        {
            var run = GetRun(text);
            run.Bold();

            return this;
        }

        public Paragraph AppendItalic(string text)
        {
            var run = GetRun(text);
            run.Italic();

            return this;
        }

        public Paragraph AppendUnderline(string text, UnderlineValues value)
        {
            var run = GetRun(text);
            run.Underline(value);

            return this;
        }

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

        public Paragraph SetAlignment(JustificationValues value)
        {
            var paragraphProperties = paragraph.GetOrCreate<ParagraphProperties>(true);
            var justification = paragraphProperties.GetOrCreate<Justification>();
            justification.Val = value;

            return this;
        }

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
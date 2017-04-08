using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

namespace DocXPlus.Models
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

        public Paragraph Alignment(JustificationValues value)
        {
            var paragraphProperties = paragraph.GetOrCreate<ParagraphProperties>(true);
            var justification = paragraphProperties.GetOrCreate<Justification>();
            justification.Val = value;

            return this;
        }

        public Paragraph Append(string text)
        {
            var run = paragraph.AppendChild(NewRun());
            run.AppendChild(new Text(text));

            return this;
        }

        public Paragraph AppendBold(string text)
        {
            var run = paragraph.AppendChild(NewRun());
            run.Bold();
            run.AppendChild(new Text(text));

            return this;
        }

        public Paragraph AppendItalic(string text)
        {
            var run = paragraph.AppendChild(NewRun());
            run.Italic();
            run.AppendChild(new Text(text));

            return this;
        }

        public Paragraph AppendUnderline(string text, UnderlineValues value)
        {
            var run = paragraph.AppendChild(NewRun());
            run.Underline(value);
            run.AppendChild(new Text(text));

            return this;
        }

        public Paragraph Bold()
        {
            if (Runs.Count() == 0)
            {
                var paragraphProperties = paragraph.GetOrCreate<ParagraphProperties>();
                var paragraphMarkRunProperties = paragraphProperties.GetOrCreate<ParagraphMarkRunProperties>();

                Bold bold = paragraphMarkRunProperties.GetOrCreate<Bold>();
                bold.Val = OnOffValue.FromBoolean(true);
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
    }
}
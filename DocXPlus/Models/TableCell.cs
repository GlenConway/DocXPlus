using System.Collections.Generic;

namespace DocXPlus.Models
{
    public class TableCell
    {
        private DocumentFormat.OpenXml.Wordprocessing.TableCell tableCell;

        private TableRow tableRow;

        internal TableCell(TableRow tableRow, DocumentFormat.OpenXml.Wordprocessing.TableCell tableCell)
        {
            this.tableRow = tableRow;
            this.tableCell = tableCell;
        }

        /// <summary>
        /// Adds a paragraph to the table cell
        /// </summary>
        /// <returns></returns>
        public Paragraph AddParagraph()
        {
            var paragraph = tableCell.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
            return new Paragraph(paragraph);
        }

        public IEnumerable<Paragraph> Paragraphs
        {
            get
            {
                var paragraphs = tableCell.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();

                var result = new List<Paragraph>();

                foreach (var paragraph in paragraphs)
                {
                    result.Add(new Paragraph(paragraph));
                }

                return result;
            }
        }
    }
}
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

namespace DocXPlus
{
    /// <summary>
    /// Represents a cell in a table
    /// </summary>
    public class TableCell
    {
        private int mergeDown;
        private int mergeRight;
        private DocumentFormat.OpenXml.Wordprocessing.TableCell tableCell;

        private TableRow tableRow;

        internal TableCell(TableRow tableRow, DocumentFormat.OpenXml.Wordprocessing.TableCell tableCell)
        {
            this.tableRow = tableRow;
            this.tableCell = tableCell;
        }

        /// <summary>
        /// Cell borders
        /// </summary>
        public Borders Borders
        {
            get
            {
                return new Borders(GetTableCellBorders());
            }
        }

        /// <summary>
        /// Merges this cell with the cells in the same column for the supplied number of rows. Does not merge the cell contents.
        /// </summary>
        public int MergeDown
        {
            get
            {
                return mergeDown;
            }
            set
            {
                if (mergeDown != value)
                {
                    mergeDown = value;

                    tableRow.MergeDown(this, value);

                    GetVerticalMerge().Val = MergedCellValues.Restart;
                }
            }
        }

        /// <summary>
        /// Merges this cell with the cells to the right. Does not merge the cell contents.
        /// </summary>
        public int MergeRight
        {
            get
            {
                return mergeRight;
            }
            set
            {
                if (mergeRight != value)
                {
                    tableRow.MergeRight(this, value);

                    mergeRight = value;

                    GetGridSpan().Val = value + 1;
                }
            }
        }

        /// <summary>
        /// All of the paragraphs in this cell.
        /// </summary>
        public Paragraph[] Paragraphs
        {
            get
            {
                var paragraphs = tableCell.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToList();

                var result = new Paragraph[paragraphs.Count()];

                for (int i = 0; i < paragraphs.Count(); i++)
                {
                    result[i] = new Paragraph(paragraphs[i]);
                }

                return result;
            }
        }

        /// <summary>
        /// Cell shading
        /// </summary>
        public Shading Shading
        {
            get
            {
                return new Shading(GetTableCellShading());
            }
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

        /// <summary>
        /// Adds a paragraph with the supplied text to the table cell
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public Paragraph AddParagraph(string text)
        {
            return AddParagraph().Append(text);
        }

        /// <summary>
        /// Sets the vertical alignment of the cell
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public TableCell SetVerticalAlignment(TableVerticalAlignmentValues value)
        {
            var tableCellVerticalAlignment = GetTableCellProperties().GetOrCreate<TableCellVerticalAlignment>();
            tableCellVerticalAlignment.Val = value;

            return this;
        }

        internal GridSpan GetGridSpan()
        {
            return GetTableCellProperties().GetOrCreate<GridSpan>();
        }

        internal TableCellBorders GetTableCellBorders()
        {
            return GetTableCellProperties().GetOrCreate<TableCellBorders>();
        }

        internal TableCellProperties GetTableCellProperties()
        {
            return tableCell.GetOrCreate<TableCellProperties>();
        }

        internal DocumentFormat.OpenXml.Wordprocessing.Shading GetTableCellShading()
        {
            return GetTableCellProperties().GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.Shading>();
        }

        internal VerticalMerge GetVerticalMerge()
        {
            return GetTableCellProperties().GetOrCreate<VerticalMerge>();
        }

        internal void RemoveFromRow()
        {
            tableCell.Remove();
        }
    }
}
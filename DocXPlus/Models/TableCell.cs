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
                    // process the rows are that are merged with this one
                    tableRow.MergeRight(this, value);

                    mergeRight = value;

                    // set the gridspan to the value plus this cell
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
        /// Gets or set the vertical alignment of the cell
        /// </summary>
        public TableVerticalAlignment VerticalAlignment
        {
            get
            {
                var tableCellVerticalAlignment = GetTableCellProperties().GetOrCreate<TableCellVerticalAlignment>();
                return Convert.ToTableVerticalAlignment(tableCellVerticalAlignment.Val);
            }
            set
            {
                var tableCellVerticalAlignment = GetTableCellProperties().GetOrCreate<TableCellVerticalAlignment>();
                tableCellVerticalAlignment.Val = Convert.ToTableVerticalAlignmentValues(value);
            }
        }

        /// <summary>
        /// The width of the cell
        /// </summary>
        public double Width
        {
            get
            {
                var tableCellWidth = tableCell.GetOrCreate<TableCellWidth>();

                if (double.TryParse(tableCellWidth.Width, out double result))
                    return result;

                return 0;
            }
            set
            {
                var tableCellWidth = tableCell.GetOrCreate<TableCellWidth>();
                tableCellWidth.Width = value.ToString();
            }
        }

        /// <summary>
        /// The width of the cell
        /// </summary>
        public TableWidthUnitValue WidthType
        {
            get
            {
                var tableCellWidth = tableCell.GetOrCreate<TableCellWidth>();
                return Convert.ToTableWidthUnitValue(tableCellWidth.Type);
            }
            set
            {
                var tableCellWidth = tableCell.GetOrCreate<TableCellWidth>();
                tableCellWidth.Type = Convert.ToTableWidthUnitValues(value);
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
        /// Adds a paragraph with the supplied text to the table cell and sets the alignment
        /// </summary>
        /// <param name="text"></param>
        /// <param name="alignment"></param>
        /// <returns></returns>
        public Paragraph AddParagraph(string text, Align alignment)
        {
            return AddParagraph().Append(text).SetAlignment(alignment);
        }

        /// <summary>
        /// Adds a Table to the document with the specified number of columns
        /// </summary>
        /// <param name="numberOfColumns"></param>
        /// <returns></returns>
        public Table AddTable(int numberOfColumns)
        {
            return tableRow.Table.Document.AddTable(numberOfColumns, AddTable());
        }

        /// <summary>
        /// Adds a Table to the document with the specified number of columns using the percent widths
        /// </summary>
        /// <param name="numberOfColumns"></param>
        /// <param name="percent"></param>
        /// <returns></returns>
        public Table AddTable(int numberOfColumns, params int[] percent)
        {
            return tableRow.Table.Document.AddTable(numberOfColumns, AddTable(), percent);
        }

        /// <summary>
        /// Adds a Table to the document with the specified number of columns using the supplied widths
        /// </summary>
        /// <param name="numberOfColumns"></param>
        /// <param name="widths">The widths of the columns in Twips</param>
        /// <returns></returns>
        public Table AddTable(int numberOfColumns, params string[] widths)
        {
            return tableRow.Table.Document.AddTable(numberOfColumns, AddTable(), widths);
        }

        /// <summary>
        /// Sets the vertical alignment of the cell
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public TableCell SetVerticalAlignment(TableVerticalAlignment value)
        {
            var tableCellVerticalAlignment = GetTableCellProperties().GetOrCreate<TableCellVerticalAlignment>();
            tableCellVerticalAlignment.Val = Convert.ToTableVerticalAlignmentValues(value);

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

        private DocumentFormat.OpenXml.Wordprocessing.Table AddTable()
        {
            var paragraph = tableCell.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Last();
            return paragraph.InsertBeforeSelf(new DocumentFormat.OpenXml.Wordprocessing.Table());
        }
    }
}
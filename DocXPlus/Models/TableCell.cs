using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Linq;

namespace DocXPlus
{
    /// <summary>
    /// Represents a cell in a table
    /// </summary>
    public class TableCell : Container, IContainer
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
        /// The width of the cell
        /// </summary>
        public override UInt32Value AvailableWidth => UInt32Value.FromUInt32(System.Convert.ToUInt32(Width));

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
        /// The margins for the table cell
        /// </summary>
        public TableCellMargin Margins
        {
            get
            {
                return new TableCellMargin(GetTableCellProperties().GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.TableCellMargin>());
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
                var tableCellVerticalAlignment = GetTableCellVerticalAlignment();
                return Convert.ToTableVerticalAlignment(tableCellVerticalAlignment.Val);
            }
            set
            {
                var tableCellVerticalAlignment = GetTableCellVerticalAlignment();
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
                var tableCellWidth = GetTableCellWidth();

                if (double.TryParse(tableCellWidth.Width, out double result))
                {
                    return result;
                }

                return 0;
            }
            set
            {
                var tableCellWidth = GetTableCellWidth();
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
                var tableCellWidth = GetTableCellWidth();
                return Convert.ToTableWidthUnitValue(tableCellWidth.Type);
            }
            set
            {
                var tableCellWidth = GetTableCellWidth();
                tableCellWidth.Type = Convert.ToTableWidthUnitValues(value);
            }
        }

        /// <summary>
        /// Sets the vertical alignment of the cell
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public TableCell SetVerticalAlignment(TableVerticalAlignment value)
        {
            var tableCellVerticalAlignment = GetTableCellVerticalAlignment();
            tableCellVerticalAlignment.Val = Convert.ToTableVerticalAlignmentValues(value);

            return this;
        }

        internal GridSpan GetGridSpan()
        {
            OpenXmlElement after = null;

            after = GetTableCellProperties().Find<TableCellWidth>();

            if (after == null)
            {
                return GetTableCellProperties().GetOrCreate<GridSpan>();
            }

            return GetTableCellProperties().GetOrCreateAfter<GridSpan>(after);
        }

        internal TableCellBorders GetTableCellBorders()
        {
            return GetTableCellProperties().GetOrCreate<TableCellBorders>();
        }

        internal TableCellProperties GetTableCellProperties()
        {
            return tableCell.GetOrCreate<TableCellProperties>(true);
        }

        internal DocumentFormat.OpenXml.Wordprocessing.Shading GetTableCellShading()
        {
            return GetTableCellProperties().GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.Shading>();
        }

        internal TableCellVerticalAlignment GetTableCellVerticalAlignment()
        {
            return GetTableCellProperties().GetOrCreate<TableCellVerticalAlignment>();
        }

        internal TableCellWidth GetTableCellWidth()
        {
            return GetTableCellProperties().GetOrCreate<TableCellWidth>(true);
        }

        internal VerticalMerge GetVerticalMerge()
        {
            OpenXmlElement after = null;

            after = GetTableCellProperties().Find<HorizontalMerge>();

            if (after == null)
            {
                after = GetTableCellProperties().Find<GridSpan>();

                if (after == null)
                {
                    after = GetTableCellProperties().Find<TableCellWidth>();

                    if (after == null)
                    {
                        return GetTableCellProperties().GetOrCreate<VerticalMerge>(true);
                    }
                }
            }

            return GetTableCellProperties().GetOrCreateAfter<VerticalMerge>(after);
        }

        internal void RemoveFromRow()
        {
            tableCell.Remove();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="contentType"></param>
        /// <returns></returns>
        protected override string AddImagePart(Stream stream, string contentType)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Adds a paragraph to the table cell
        /// </summary>
        /// <returns></returns>
        protected override DocumentFormat.OpenXml.Wordprocessing.Paragraph NewParagraph()
        {
            return tableCell.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
        }

        /// <summary>
        /// Adds a new table before the last paragraph in the cell
        /// </summary>
        /// <returns></returns>
        protected override DocumentFormat.OpenXml.Wordprocessing.Table NewTable()
        {
            var paragraph = tableCell.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Last();
            return paragraph.InsertBeforeSelf(new DocumentFormat.OpenXml.Wordprocessing.Table());
        }
    }
}
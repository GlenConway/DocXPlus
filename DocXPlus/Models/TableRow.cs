using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;

namespace DocXPlus
{
    /// <summary>
    /// Represents a row in a table
    /// </summary>
    public class TableRow
    {
        private TableCell[] cells;
        private Table table;
        private DocumentFormat.OpenXml.Wordprocessing.TableRow tableRow;

        internal TableRow(Table table, DocumentFormat.OpenXml.Wordprocessing.TableRow tableRow)
        {
            this.table = table;

            this.tableRow = tableRow;

            AddCells();
        }

        /// <summary>
        /// Allow the row to break across pages
        /// </summary>
        public bool BreakAcrossPages
        {
            get
            {
                return GetCantSplit().Val == OnOffOnlyValues.Off;
            }
            set
            {
                GetCantSplit().Val = (value ? OnOffOnlyValues.Off : OnOffOnlyValues.On);
            }
        }

        /// <summary>
        /// The cells in the row
        /// </summary>
        public TableCell[] Cells => cells;

        /// <summary>
        /// Gets or sets if the row is a header
        /// </summary>
        public bool HeaderRow
        {
            get
            {
                return GetTableHeader().Val == OnOffOnlyValues.On;
            }
            set
            {
                GetTableHeader().Val = value.ToOnOffOnlyValues();
            }
        }

        /// <summary>
        /// Height of the row in Twips
        /// </summary>
        public UInt32Value Height
        {
            get
            {
                return GetTableRowHeights().Val;
            }
            set
            {
                GetTableRowHeights().Val = value;
            }
        }

        /// <summary>
        /// The type of height
        /// </summary>
        public HeightRuleValues HeightType
        {
            get
            {
                return GetTableRowHeights().HeightType;
            }
            set
            {
                GetTableRowHeights().HeightType = value;
            }
        }

        /// <summary>
        /// The table that contains the row
        /// </summary>
        public Table Table => table;

        /// <summary>
        ///
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public TableCell this[int index]
        {
            get
            {
                return cells[index];
            }
        }

        /// <summary>
        /// Sets the Top, Bottom, Left and Right borders for every cell in the row.
        /// </summary>
        /// <param name="size">The size of the border in Twips</param>
        /// <param name="value"></param>
        /// <param name="color"></param>
        public void SetBorders(UInt32Value size, BorderValue value, string color = "auto")
        {
            foreach (var cell in Cells)
            {
                cell.Borders.Set(size, value, color);
            }
        }

        /// <summary>
        /// Sets the Top, Bottom, Left and Right shading for every cell in the row
        /// </summary>
        /// <param name="value"></param>
        /// <param name="fill">Hex fill color</param>
        /// <param name="color"></param>
        public void SetShading(ShadingPatternValue value, string fill, string color = "auto")
        {
            foreach (var cell in Cells)
            {
                cell.Shading.Set(value, fill, color);
            }
        }

        internal CantSplit GetCantSplit()
        {
            return GetTableRowProperties().GetOrCreate<CantSplit>();
        }

        internal TableHeader GetTableHeader()
        {
            return GetTableRowProperties().GetOrCreate<TableHeader>();
        }

        internal TableRowHeight GetTableRowHeights()
        {
            return GetTableRowProperties().GetOrCreate<TableRowHeight>();
        }

        internal TableRowProperties GetTableRowProperties()
        {
            return tableRow.GetOrCreate<TableRowProperties>();
        }

        internal void MergeDown(TableCell tableCell, int value)
        {
            var index = Cells.ToList().IndexOf(tableCell);

            table.MergeDown(this, index, value);
        }

        internal void MergeRight(TableCell tableCell, int value)
        {
            if (value == 0)
            {
                throw new ArgumentException("Value must be greater than zero. Cannot merge a cell with itself.");
            }

            if (value >= Cells.Count())
            {
                throw new ArgumentOutOfRangeException(nameof(value), $"Value {value} must be less than {Cells.Count()}");
            }

            var index = Cells.ToList().IndexOf(tableCell);

            if (value + index >= Cells.Count())
                throw new ArgumentOutOfRangeException(nameof(value), $"Value {value} must be less than {Cells.Count() - index}");

            for (int i = 1; i <= value; i++)
            {
                Cells[i].RemoveFromRow();
            }
        }

        private void AddCells()
        {
            cells = new TableCell[table.NumberOfColumns];

            for (int i = 0; i < table.NumberOfColumns; i++)
            {
                var tableCell = tableRow.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.TableCell());

                var tableCellWidth = tableCell.GetOrCreate<TableCellWidth>();
                tableCellWidth.Width = table.ColumnWidths[i];
                tableCellWidth.Type = TableWidthUnitValues.Dxa;

                var cell = new TableCell(this, tableCell);
                cell.AddParagraph();

                cells[i] = cell;
            }
        }
    }
}
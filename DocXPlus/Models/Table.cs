using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DocXPlus
{
    /// <summary>
    /// Represents a table in the document
    /// </summary>
    public class Table
    {
        private string[] columnWidths;
        private IContainer container;
        private int numberOfColumns;
        private IList<TableRow> rows;
        private DocumentFormat.OpenXml.Wordprocessing.Table table;
        private TableLook tableLook;

        internal Table(DocumentFormat.OpenXml.Wordprocessing.Table table, IContainer container)
        {
            this.table = table;

            this.container = container;

            BuildColumnWidths();
            BuildRows();
        }

        internal Table(DocumentFormat.OpenXml.Wordprocessing.Table table, int numberOfColumns, IContainer container) : this(table, container)
        {
            this.numberOfColumns = numberOfColumns;

            AddGrid();
        }

        internal Table(DocumentFormat.OpenXml.Wordprocessing.Table table, int numberOfColumns, IContainer container, params int[] percent) : this(table, container)
        {
            this.numberOfColumns = numberOfColumns;

            AddGrid(percent);
        }

        internal Table(DocumentFormat.OpenXml.Wordprocessing.Table table, int numberOfColumns, IContainer container, params string[] widths) : this(table, container)
        {
            this.numberOfColumns = numberOfColumns;

            AddGrid(widths);
        }

        /// <summary>
        /// The default margins for table cells
        /// </summary>
        public TableCellMarginDefault DefaultMargins
        {
            get
            {
                return new TableCellMarginDefault(GetTableProperties().GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.TableCellMarginDefault>());
            }
        }

        /// <summary>
        /// Gets the number of columns in the table
        /// </summary>
        public int NumberOfColumns => numberOfColumns;

        /// <summary>
        /// Gets the rows in the table
        /// </summary>
        public IEnumerable<TableRow> Rows => rows;

        /// <summary>
        /// Gets or sets the style of the table
        /// </summary>
        public string TableStyle
        {
            get
            {
                var tableStyle = TableProperties.GetOrCreate<TableStyle>();
                return tableStyle.Val;
            }
            set
            {
                var tableStyle = TableProperties.GetOrCreate<TableStyle>();
                tableStyle.Val = value;
            }
        }

        /// <summary>
        /// Gets or sets the width of the table in Twips
        /// </summary>
        public string Width
        {
            get
            {
                var tableWidth = TableProperties.GetOrCreate<TableWidth>();
                return tableWidth.Width;
            }
            set
            {
                var tableWidth = TableProperties.GetOrCreate<TableWidth>();
                tableWidth.Width = value;
            }
        }

        /// <summary>
        /// Gets or sets the width type for the table
        /// </summary>
        public TableWidthUnitValues WidthType
        {
            get
            {
                var tableWidth = TableProperties.GetOrCreate<TableWidth>();
                return tableWidth.Type;
            }
            set
            {
                var tableWidth = TableProperties.GetOrCreate<TableWidth>();
                tableWidth.Type = value;
            }
        }

        internal string[] ColumnWidths => columnWidths;

        internal IContainer Document => container;

        internal TableLook TableLook
        {
            get
            {
                if (tableLook == null)
                {
                    tableLook = new TableLook(TableProperties);
                }

                return tableLook;
            }
        }

        internal TableProperties TableProperties => table.GetOrCreate<TableProperties>();

        /// <summary>
        /// Adds a row to the table. The row will have the same number of cells as the number of columns in the table.
        /// Each cell will have an empty paragraph
        /// </summary>
        /// <returns></returns>
        public TableRow AddRow()
        {
            var tableRow = table.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.TableRow());

            var result = new TableRow(this, tableRow)
            {
                HeaderRow = false
            };

            if (rows == null)
            {
                rows = new List<TableRow>();
            }

            rows.Add(result);

            return result;
        }

        internal void MergeDown(TableRow tableRow, int cellIndex, int value)
        {
            if (value == 0)
            {
                throw new ArgumentException("Value must be greater than zero. Cannot merge a cell with itself.");
            }

            if (value >= Rows.Count())
            {
                throw new ArgumentOutOfRangeException(nameof(value), $"Value {value} must be less than {Rows.Count()}");
            }

            var rows = Rows.ToList();
            var rowIndex = rows.IndexOf(tableRow);

            for (int i = 1; i <= value; i++)
            {
                rows[i].Cells[cellIndex].GetVerticalMerge().Val = MergedCellValues.Continue;
            }
        }

        private void AddGrid()
        {
            var tableGrid = table.AppendChild(new TableGrid());

            var columnWidth = container.AvailableWidth / NumberOfColumns;

            columnWidths = new string[NumberOfColumns];

            for (int i = 0; i < NumberOfColumns; i++)
            {
                var gridColumn = tableGrid.AppendChild(new GridColumn());
                gridColumn.Width = columnWidth.ToString();

                columnWidths[i] = columnWidth.ToString();
            }
        }

        private void AddGrid(params int[] percent)
        {
            if (percent.Sum() != 100)
                throw new ArgumentException("Widths must add up to 100%");

            if (percent.Count() != NumberOfColumns)
                throw new ArgumentException("Widths must equal the number of columns");

            var tableGrid = table.AppendChild(new TableGrid());

            var width = container.AvailableWidth;

            columnWidths = new string[NumberOfColumns];

            for (int i = 0; i < NumberOfColumns; i++)
            {
                var columnWidth = ((double)percent[i] / 100);

                columnWidths[i] = (width * columnWidth).ToString();
            }

            for (int i = 0; i < NumberOfColumns; i++)
            {
                var gridColumn = tableGrid.AppendChild(new GridColumn());
                gridColumn.Width = columnWidths[i];
            }
        }

        private void AddGrid(params string[] widths)
        {
            if (widths.Count() != NumberOfColumns)
                throw new ArgumentException("Widths must equal the number of columns");

            var tableGrid = table.AppendChild(new TableGrid());

            columnWidths = new string[NumberOfColumns];

            for (int i = 0; i < NumberOfColumns; i++)
            {
                columnWidths[i] = widths[i];
            }

            for (int i = 0; i < NumberOfColumns; i++)
            {
                var gridColumn = tableGrid.AppendChild(new GridColumn());
                gridColumn.Width = columnWidths[i];
            }
        }

        private void BuildColumnWidths()
        {
            if (!table.Has<TableGrid>())
            {
                return;
            }

            var tableGrid = table.GetOrCreate<TableGrid>();

            var grids = tableGrid.Descendants<GridColumn>().ToArray();

            numberOfColumns = grids.Count();
            columnWidths = new string[numberOfColumns];

            for (int i = 0; i < numberOfColumns; i++)
            {
                columnWidths[i] = grids[i].Width;
            }
        }

        private void BuildRows()
        {
            var tableRows = table.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableRow>();

            rows = new List<TableRow>();

            foreach (var row in tableRows)
            {
                rows.Add(new TableRow(this, row));
            }
        }

        private TableProperties GetTableProperties()
        {
            return table.GetOrCreate<TableProperties>();
        }
    }
}
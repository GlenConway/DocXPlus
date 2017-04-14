using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DocXPlus
{
    public class Table
    {
        private string[] columnWidths;
        private DocX document;
        private int numberOfColumns;
        private IList<TableRow> rows;
        private DocumentFormat.OpenXml.Wordprocessing.Table table;
        private TableLook tableLook;

        internal Table(DocumentFormat.OpenXml.Wordprocessing.Table table, int numberOfColumns, DocX document)
        {
            this.table = table;
            this.numberOfColumns = numberOfColumns;
            this.document = document;

            AddGrid();
        }

        public int NumberOfColumns => numberOfColumns;

        public IEnumerable<TableRow> Rows => rows;

        public TableLook TableLook
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

            var width = document.PageWidth.Value - document.PageMargins.RightAndLeft.Value;
            var columnWidth = width / NumberOfColumns;

            columnWidths = new string[NumberOfColumns];

            for (int i = 0; i < NumberOfColumns; i++)
            {
                var gridColumn = tableGrid.AppendChild(new GridColumn());
                gridColumn.Width = columnWidth.ToString();

                columnWidths[i] = columnWidth.ToString();
            }
        }
    }
}
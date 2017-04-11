using DocumentFormat.OpenXml.Wordprocessing;

namespace DocXPlus.Models
{
    public class Table
    {
        private string[] columnWidths;
        private DocX document;
        private int numberOfColumns;
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

        public TableRow AddRow()
        {
            var tableRow = table.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.TableRow());
            var tableRowProperties = tableRow.GetOrCreate<TableRowProperties>();

            var result = new TableRow(this, tableRow);
            return result;
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
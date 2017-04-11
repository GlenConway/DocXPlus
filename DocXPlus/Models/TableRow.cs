using System.Collections.Generic;

namespace DocXPlus.Models
{
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

        public Table Table => table;

        public TableCell this[int index]
        {
            get
            {
                return cells[index];
            }
        }

        private void AddCells()
        {
            cells = new TableCell[table.NumberOfColumns];

            for (int i = 0; i < table.NumberOfColumns; i++)
            {
                var tableCell = tableRow.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.TableCell());

                var tableCellWidth = tableCell.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.TableCellWidth>();
                tableCellWidth.Width = table.ColumnWidths[i];
                tableCellWidth.Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Dxa;

                var cell = new TableCell(this, tableCell);
                cell.AddParagraph();

                cells[i] = cell;
            }
        }
    }
}
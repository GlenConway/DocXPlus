namespace DocXPlus
{
    /// <summary>
    /// Defines table indentation values
    /// </summary>
    public class TableIndentation
    {
        private DocumentFormat.OpenXml.Wordprocessing.TableIndentation tableIndentation;

        /// <summary>
        ///
        /// </summary>
        /// <param name="tableIndentation"></param>
        public TableIndentation(DocumentFormat.OpenXml.Wordprocessing.TableIndentation tableIndentation)
        {
            this.tableIndentation = tableIndentation;
        }

        /// <summary>
        /// The width of the cell
        /// </summary>
        public int Width
        {
            get
            {
                return tableIndentation.Width;
            }
            set
            {
                tableIndentation.Width = value;
            }
        }

        /// <summary>
        /// The width of the cell
        /// </summary>
        public TableWidthUnitValue WidthType
        {
            get
            {
                return Convert.ToTableWidthUnitValue(tableIndentation.Type);
            }
            set
            {
                tableIndentation.Type = Convert.ToTableWidthUnitValues(value);
            }
        }
    }
}
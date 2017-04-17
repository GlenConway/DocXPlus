namespace DocXPlus
{
    /// <summary>
    /// Represents a table with
    /// </summary>
    public class TableWidthDxaNilType
    {
        private DocumentFormat.OpenXml.Wordprocessing.TableWidthDxaNilType tableWidthType;

        internal TableWidthDxaNilType(DocumentFormat.OpenXml.Wordprocessing.TableWidthDxaNilType tableWidthType)
        {
            this.tableWidthType = tableWidthType;
        }

        /// <summary>
        /// The width in Twips
        /// </summary>
        public TableWidthValue Type
        {
            get
            {
                return Convert.ToTableWidthValue(tableWidthType.Type);
            }
            set
            {
                tableWidthType.Type = Convert.ToTableWidthValues(value);
            }
        }

        /// <summary>
        /// The width in Twips
        /// </summary>
        public System.Int16 Width
        {
            get
            {
                return tableWidthType.Width;
            }
            set
            {
                tableWidthType.Width = value;
            }
        }
    }
}
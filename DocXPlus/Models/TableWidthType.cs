namespace DocXPlus
{
    /// <summary>
    /// Represents a table with
    /// </summary>
    public class TableWidthType
    {
        private DocumentFormat.OpenXml.Wordprocessing.TableWidthType tableWidthType;

        internal TableWidthType(DocumentFormat.OpenXml.Wordprocessing.TableWidthType tableWidthType)
        {
            this.tableWidthType = tableWidthType;
        }

        /// <summary>
        /// The width in Twips
        /// </summary>
        public TableWidthUnitValue Type
        {
            get
            {
                return Convert.ToTableWidthUnitValue(tableWidthType.Type);
            }
            set
            {
                tableWidthType.Type = Convert.ToTableWidthUnitValues(value);
            }
        }

        /// <summary>
        /// The width in Twips
        /// </summary>
        public string Width
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

        /// <summary>
        /// Sets the width in centimetres
        /// </summary>
        /// <param name="value"></param>
        public void SetCm(double value)
        {
            Width = Units.CMToTwips(value).ToString();
            Type = TableWidthUnitValue.Dxa;
        }

        /// <summary>
        /// Sets the width in inches
        /// </summary>
        /// <param name="value"></param>
        public void SetInch(double value)
        {
            Width = Units.InchToTwips(value).ToString();
            Type = TableWidthUnitValue.Dxa;
        }
    }
}
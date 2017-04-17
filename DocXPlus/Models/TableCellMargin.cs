namespace DocXPlus
{
    /// <summary>
    /// Represents a table cell margin
    /// </summary>
    public class TableCellMargin
    {
        private DocumentFormat.OpenXml.Wordprocessing.TableCellMargin tableCellMargin;
        
        internal TableCellMargin(DocumentFormat.OpenXml.Wordprocessing.TableCellMargin tableCellMargin)
        {
            this.tableCellMargin = tableCellMargin;
        }

        /// <summary>
        /// Bottom margin
        /// </summary>
        public TableWidthType BottomMargin
        {
            get
            {
                return new TableWidthType(tableCellMargin.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.BottomMargin>());
            }
        }

        /// <summary>
        /// End margin
        /// </summary>
        public TableWidthType EndMargin
        {
            get
            {
                return new TableWidthType(tableCellMargin.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.EndMargin>());
            }
        }

        /// <summary>
        /// Left margin
        /// </summary>
        public TableWidthType LeftMargin
        {
            get
            {
                return new TableWidthType(tableCellMargin.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.LeftMargin>());
            }
        }

        /// <summary>
        /// Right margin
        /// </summary>
        public TableWidthType RightMargin
        {
            get
            {
                return new TableWidthType(tableCellMargin.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.RightMargin>());
            }
        }

        /// <summary>
        /// Start margin
        /// </summary>
        public TableWidthType StartMargin
        {
            get
            {
                return new TableWidthType(tableCellMargin.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.StartMargin>());
            }
        }

        /// <summary>
        /// Top margin
        /// </summary>
        public TableWidthType TopMargin
        {
            get
            {
                return new TableWidthType(tableCellMargin.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.TopMargin>());
            }
        }
    }
}
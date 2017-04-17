namespace DocXPlus
{
    /// <summary>
    /// Represents a table cell margin
    /// </summary>
    public class TableCellMarginDefault
    {
        private DocumentFormat.OpenXml.Wordprocessing.TableCellMarginDefault tableCellMarginDefault;

        internal TableCellMarginDefault(DocumentFormat.OpenXml.Wordprocessing.TableCellMarginDefault tableCellMarginDefault)
        {
            this.tableCellMarginDefault = tableCellMarginDefault;
        }

        /// <summary>
        /// Bottom margin
        /// </summary>
        public TableWidthType BottomMargin
        {
            get
            {
                return new TableWidthType(tableCellMarginDefault.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.BottomMargin>());
            }
        }

        /// <summary>
        /// End margin
        /// </summary>
        public TableWidthType EndMargin
        {
            get
            {
                return new TableWidthType(tableCellMarginDefault.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.EndMargin>());
            }
        }

        /// <summary>
        /// Left margin
        /// </summary>
        public TableWidthDxaNilType LeftMargin
        {
            get
            {
                return new TableWidthDxaNilType(tableCellMarginDefault.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.TableCellLeftMargin>());
            }
        }

        /// <summary>
        /// Right margin
        /// </summary>
        public TableWidthDxaNilType RightMargin
        {
            get
            {
                return new TableWidthDxaNilType(tableCellMarginDefault.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.TableCellRightMargin>());
            }
        }

        /// <summary>
        /// Start margin
        /// </summary>
        public TableWidthType StartMargin
        {
            get
            {
                return new TableWidthType(tableCellMarginDefault.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.StartMargin>());
            }
        }

        /// <summary>
        /// Top margin
        /// </summary>
        public TableWidthType TopMargin
        {
            get
            {
                return new TableWidthType(tableCellMarginDefault.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.TopMargin>());
            }
        }
    }
}
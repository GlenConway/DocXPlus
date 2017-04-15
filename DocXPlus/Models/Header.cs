using DocumentFormat.OpenXml.Wordprocessing;

namespace DocXPlus
{
    /// <summary>
    /// Represents a header in the document
    /// </summary>
    public class Header
    {
        private DocX document;
        private DocumentFormat.OpenXml.Wordprocessing.Header header;
        private HeaderFooterValues type;

        internal Header(DocumentFormat.OpenXml.Wordprocessing.Header header, DocX document, HeaderFooterValues type)
        {
            this.header = header;
            this.document = document;
            this.type = type;
        }

        /// <summary>
        /// The type of header
        /// </summary>
        public HeaderFooterValues Type
        {
            get
            {
                return type;
            }
        }

        /// <summary>
        /// Adds a paragraph to the header
        /// </summary>
        /// <returns></returns>
        public Paragraph AddParagraph()
        {
            var paragraph = header.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
            return new Paragraph(paragraph);
        }

        /// <summary>
        /// Adds a table with the specified number of columns. Columns widths are evenly distributed.
        /// </summary>
        /// <param name="numberOfColumns"></param>
        /// <returns></returns>
        public Table AddTable(int numberOfColumns)
        {
            var table = header.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Table());

            return document.AddTable(numberOfColumns, table);
        }

        /// <summary>
        /// Adds a table with the specified number of columns. Column widths are calculated based on the supplied percent values.
        /// </summary>
        /// <param name="numberOfColumns"></param>
        /// <param name="percent"></param>
        /// <returns></returns>
        public Table AddTable(int numberOfColumns, params int[] percent)
        {
            var table = header.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Table());

            return document.AddTable(numberOfColumns, table, percent);
        }

        /// <summary>
        /// Adds a table with the specified number of columns. Column widths are based on the supplied width values.
        /// </summary>
        /// <param name="numberOfColumns"></param>
        /// <param name="widths">The width of the columns in Twips</param>
        /// <returns></returns>
        public Table AddTable(int numberOfColumns, params string[] widths)
        {
            var table = header.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Table());

            return document.AddTable(numberOfColumns, table, widths);
        }

        internal void Save()
        {
            header.Save();
        }
    }
}
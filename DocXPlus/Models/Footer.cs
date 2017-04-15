using DocumentFormat.OpenXml.Wordprocessing;

namespace DocXPlus
{
    /// <summary>
    /// Represents a footer in the document
    /// </summary>
    public class Footer
    {
        private DocX document;
        private DocumentFormat.OpenXml.Wordprocessing.Footer footer;
        private HeaderFooterValues type;

        internal Footer(DocumentFormat.OpenXml.Wordprocessing.Footer footer, DocX document, HeaderFooterValues type)
        {
            this.footer = footer;
            this.document = document;
            this.type = type;
        }

        /// <summary>
        /// The type of footer
        /// </summary>
        public HeaderFooterValues Type
        {
            get
            {
                return type;
            }
        }

        /// <summary>
        /// Adds a paragraph to the footer
        /// </summary>
        /// <returns></returns>
        public Paragraph AddParagraph()
        {
            var paragraph = footer.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
            return new Paragraph(paragraph);
        }

        /// <summary>
        /// Adds a table with the specified number of columns. Columns widths are evenly distributed.
        /// </summary>
        /// <param name="numberOfColumns"></param>
        /// <returns></returns>
        public Table AddTable(int numberOfColumns)
        {
            var table = footer.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Table());

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
            var table = footer.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Table());

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
            var table = footer.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Table());

            return document.AddTable(numberOfColumns, table, widths);
        }

        internal void Save()
        {
            footer.Save();
        }
    }
}
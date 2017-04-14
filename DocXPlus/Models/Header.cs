using DocumentFormat.OpenXml.Wordprocessing;

namespace DocXPlus
{
    public class Header
    {
        private DocX document;
        private DocumentFormat.OpenXml.Wordprocessing.Header header;
        private HeaderFooterValues type;

        public Header(DocumentFormat.OpenXml.Wordprocessing.Header header, DocX document, HeaderFooterValues type)
        {
            this.header = header;
            this.document = document;
            this.type = type;
        }

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

        public Table AddTable(int numberOfColumns)
        {
            var table = header.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Table());

            return document.AddTable(numberOfColumns, table);
        }

        public void Save()
        {
            header.Save();
        }
    }
}
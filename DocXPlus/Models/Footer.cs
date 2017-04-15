using DocumentFormat.OpenXml.Wordprocessing;

namespace DocXPlus
{
    public class Footer
    {
        private DocX document;
        private DocumentFormat.OpenXml.Wordprocessing.Footer footer;
        private HeaderFooterValues type;

        public Footer(DocumentFormat.OpenXml.Wordprocessing.Footer footer, DocX document, HeaderFooterValues type)
        {
            this.footer = footer;
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

        public Paragraph AddParagraph()
        {
            var paragraph = footer.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
            return new Paragraph(paragraph);
        }

        public Table AddTable(int numberOfColumns)
        {
            var table = footer.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Table());

            return document.AddTable(numberOfColumns, table);
        }

        public Table AddTable(int numberOfColumns, params int[] percent)
        {
            var table = footer.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Table());

            return document.AddTable(numberOfColumns, table, percent);
        }

        public void Save()
        {
            footer.Save();
        }
    }
}
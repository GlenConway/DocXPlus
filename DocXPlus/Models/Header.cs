namespace DocXPlus.Models
{
    public class Header
    {
        private DocumentFormat.OpenXml.Wordprocessing.Header header;

        public Header(DocumentFormat.OpenXml.Wordprocessing.Header header)
        {
            this.header = header;
        }

        public Paragraph AddParagraph()
        {
            var paragraph = header.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
            return new Paragraph(paragraph);
        }

        public void Save()
        {
            header.Save();
        }
    }
}
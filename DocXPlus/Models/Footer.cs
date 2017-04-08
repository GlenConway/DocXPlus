namespace DocXPlus.Models
{
    public class Footer
    {
        private DocumentFormat.OpenXml.Wordprocessing.Footer footer;

        public Footer(DocumentFormat.OpenXml.Wordprocessing.Footer footer)
        {
            this.footer = footer;
        }

        public Paragraph AddParagraph()
        {
            var paragraph = footer.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
            return new Paragraph(paragraph);
        }

        public void Save()
        {
            footer.Save();
        }
    }
}
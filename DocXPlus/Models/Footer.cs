using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace DocXPlus
{
    /// <summary>
    /// Represents a footer in the document
    /// </summary>
    public class Footer : Container
    {
        private DocX document;
        private DocumentFormat.OpenXml.Wordprocessing.Footer footer;
        private FooterPart footerPart;
        private HeaderFooterValues type;

        internal Footer(FooterPart part, DocX document, HeaderFooterValues type)
        {
            footer = part.Footer;
            footerPart = part;

            this.document = document;
            this.type = type;
        }

        /// <summary>
        /// Returns the available width of the document
        /// </summary>
        public override UInt32Value AvailableWidth => document.AvailableWidth;

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

        internal void Save()
        {
            footer.Save();
        }

        /// <summary>
        /// Adds an image part to the footer and returns the part ID
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="contentType"></param>
        /// <returns></returns>
        protected override string AddImagePart(Stream stream, string contentType)
        {
            var imagePart = footerPart.AddImagePart(contentType);
            imagePart.FeedData(stream);

            return footerPart.GetIdOfPart(imagePart);
        }

        /// <summary>
        /// Adds a paragraph to the footer
        /// </summary>
        /// <returns></returns>
        protected override DocumentFormat.OpenXml.Wordprocessing.Paragraph NewParagraph()
        {
            return footer.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
        }

        /// <summary>
        /// Appends a table to the footer
        /// </summary>
        /// <returns></returns>
        protected override DocumentFormat.OpenXml.Wordprocessing.Table NewTable()
        {
            return footer.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Table());
        }
    }
}
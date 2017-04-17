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

        /// <summary>
        /// Adds an image to the footer which can then be added to a paragraph
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="x">The width of the image in English Metric Units (EMU)</param>
        /// <param name="y">The height of the image in English Metric Units (EMU)</param>
        /// <returns></returns>
        public Drawing AddImage(string fileName, Int64Value x, Int64Value y)
        {
            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                return AddImage(stream, DocX.FileNameContentType(fileName), x, y);
            }
        }

        /// <summary>
        /// Adds an image to the footer which can then be added to a paragraph
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="contentType"></param>
        /// <param name="width">The width of the image in English Metric Units (EMU)</param>
        /// <param name="height">The height of the image in English Metric Units (EMU)</param>
        /// <returns></returns>
        public Drawing AddImage(Stream stream, string contentType, Int64Value width, Int64Value height)
        {
            ImagePart imagePart = footerPart.AddImagePart(contentType);
            imagePart.FeedData(stream);

            return DocX.CreateDrawing(footerPart.GetIdOfPart(imagePart), width, height);
        }

        /// <summary>
        /// Adds an image to the footer which can then be added to a paragraph
        /// </summary>
        /// <param name="data"></param>
        /// <param name="contentType"></param>
        /// <param name="width">The width of the image in English Metric Units (EMU)</param>
        /// <param name="height">The height of the image in English Metric Units (EMU)</param>
        /// <returns></returns>
        public Drawing AddImage(byte[] data, string contentType, Int64Value width, Int64Value height)
        {
            using (var stream = new MemoryStream(data))
            {
                return AddImage(stream, contentType, width, height);
            }
        }

        internal void Save()
        {
            footer.Save();
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
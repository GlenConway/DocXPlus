using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System;

namespace DocXPlus
{
    /// <summary>
    /// Represents a header in the document
    /// </summary>
    public class Header : Container, IContainer
    {
        private DocX document;
        private DocumentFormat.OpenXml.Wordprocessing.Header header;
        private HeaderPart headerPart;
        private HeaderFooterValues type;

        internal Header(HeaderPart part, DocX document, HeaderFooterValues type)
        {
            headerPart = part;
            header = part.Header;

            this.document = document;
            this.type = type;
        }

        /// <summary>
        /// Returns the available width of the document
        /// </summary>
        public override UInt32Value AvailableWidth => document.AvailableWidth;

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
        /// Adds an image part to the header and returns the part ID
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="contentType"></param>
        /// <returns></returns>
        protected override string AddImagePart(Stream stream, string contentType)
        {
            var imagePart = headerPart.AddImagePart(contentType);
            imagePart.FeedData(stream);

            return headerPart.GetIdOfPart(imagePart);
        }

        internal void Save()
        {
            header.Save();
        }

        /// <summary>
        /// Adds a paragraph to the header
        /// </summary>
        /// <returns></returns>
        protected override DocumentFormat.OpenXml.Wordprocessing.Paragraph NewParagraph()
        {
            return header.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
        }

        /// <summary>
        /// Appends a table to the header
        /// </summary>
        /// <returns></returns>
        protected override DocumentFormat.OpenXml.Wordprocessing.Table NewTable()
        {
            return header.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Table());
        }
    }
}
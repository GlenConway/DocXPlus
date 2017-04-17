using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace DocXPlus
{
    /// <summary>
    /// Represents a header in the document
    /// </summary>
    public class Header
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
            ImagePart imagePart = headerPart.AddImagePart(contentType);
            imagePart.FeedData(stream);

            return DocX.CreateDrawing(headerPart.GetIdOfPart(imagePart), width, height);
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
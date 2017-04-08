using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace DocXPlus
{
    public class DocX
    {
        private WordprocessingDocument document;

        internal Body Body
        {
            get
            {
                return document.MainDocumentPart.Document.Body;
            }
        }

        public static DocX Create(string path, WordprocessingDocumentType type)
        {
            return Create(path, type, false);
        }

        public static DocX Create(string path, WordprocessingDocumentType type, bool autoSave)
        {
            var docX = new DocX();
            docX.Create(WordprocessingDocument.Create(path, type, autoSave));

            return docX;
        }

        public static DocX Create(Stream stream, WordprocessingDocumentType type, bool autoSave)
        {
            var docX = new DocX();
            docX.Create(WordprocessingDocument.Create(stream, type, autoSave));

            return docX;
        }

        public Models.Paragraph AddParagraph()
        {
            var paragraph = Body.AppendChild(new Paragraph());
            return new Models.Paragraph(paragraph);
        }

        public void Close()
        {
            document.MainDocumentPart.Document.Save();
            document.Close();
        }

        public void Save()
        {
            document.MainDocumentPart.Document.Save();
            document.Save();
        }

        internal void Create(WordprocessingDocument doc)
        {
            this.document = doc;

            // Add a main document part.
            MainDocumentPart mainPart = doc.AddMainDocumentPart();

            mainPart.Document = new Document();
            mainPart.Document.AppendChild(new Body());
        }
    }
}
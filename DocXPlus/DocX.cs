using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace DocXPlus
{
    public class DocX
    {
        private WordprocessingDocument document;

        public static DocX Create(string path, WordprocessingDocumentType type)
        {
            return DocX.Create(path, type, false);
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

        internal void Create(WordprocessingDocument doc)
        {
            document = doc;

            // Add a main document part. 
            MainDocumentPart mainPart = document.AddMainDocumentPart();

            // Create the document structure
            mainPart.Document = new Document();
        }
    }
}
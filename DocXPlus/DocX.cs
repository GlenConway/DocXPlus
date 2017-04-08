using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace DocXPlus
{
    public class DocX
    {
        private WordprocessingDocument document;

        private Models.Footer footer;
        private Models.Header header;

        internal Body Body
        {
            get
            {
                return MainDocumentPart.Document.Body;
            }
        }

        internal MainDocumentPart MainDocumentPart
        {
            get
            {
                return document.MainDocumentPart;
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

        public Models.Footer AddFooter()
        {
            MainDocumentPart.DeleteParts(document.MainDocumentPart.FooterParts);

            var part = MainDocumentPart.AddNewPart<FooterPart>();

            var id = MainDocumentPart.GetIdOfPart(part);

            GenerateFooterPartContent(part);

            var sectionProperty = Body.GetOrCreate<SectionProperties>();

            sectionProperty.RemoveAllChildren<FooterReference>();

            sectionProperty.PrependChild(new FooterReference() { Id = id });

            footer = new Models.Footer(part.Footer);

            return footer;
        }

        public Models.Header AddHeader()
        {
            MainDocumentPart.DeleteParts(document.MainDocumentPart.HeaderParts);

            var part = MainDocumentPart.AddNewPart<HeaderPart>();

            var id = MainDocumentPart.GetIdOfPart(part);

            GenerateHeaderPartContent(part);

            var sectionProperty = Body.GetOrCreate<SectionProperties>();

            sectionProperty.RemoveAllChildren<HeaderReference>();

            sectionProperty.PrependChild(new HeaderReference() { Id = id });

            header = new Models.Header(part.Header);

            return header;
        }

        public Models.Paragraph AddParagraph()
        {
            var paragraph = Body.AppendChild(new Paragraph());
            return new Models.Paragraph(paragraph);
        }

        public void Close()
        {
            Save();

            document.Close();
        }

        public void Save()
        {
            document.MainDocumentPart.Document.Save();

            if (header != null)
            {
                header.Save();
            }

            if (footer != null)
            {
                footer.Save();
            }

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

        internal void GenerateFooterPartContent(FooterPart part)
        {
            var footer = new Footer() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };

            footer.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            footer.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footer.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footer.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footer.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footer.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footer.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            footer.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footer.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footer.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footer.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            footer.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            footer.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            footer.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            footer.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            part.Footer = footer;
        }

        internal void GenerateHeaderPartContent(HeaderPart part)
        {
            var header = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };

            header.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            part.Header = header;
        }
    }
}
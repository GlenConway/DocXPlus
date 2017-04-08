using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Linq;

namespace DocXPlus
{
    public class DocX
    {
        private WordprocessingDocument document;

        private Models.Footer footer;
        private Models.Header header;

        internal static Int32Value Inch
        {
            get
            {
                return new Int32Value(1440);
            }
        }

        internal static UInt32Value UHalfInch
        {
            get
            {
                return new UInt32Value((uint)720);
            }
        }

        internal static UInt32Value UInch
        {
            get
            {
                return new UInt32Value((uint)1440);
            }
        }

        internal static UInt32Value UZero
        {
            get
            {
                return new UInt32Value((uint)0);
            }
        }

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

        public DocX SetOrientation(PageOrientationValues value)
        {
            bool documentChanged = false;

            var sectionProperty = Body.GetOrCreate<SectionProperties>();

            bool pageOrientationChanged = false;

            PageSize pageSize = sectionProperty.Descendants<PageSize>().FirstOrDefault();

            if (pageSize != null)
            {
                // No Orient property? Create it now. Otherwise, just
                // set its value. Assume that the default orientation
                // is Portrait.
                if (pageSize.Orient == null)
                {
                    // Need to create the attribute. You do not need to
                    // create the Orient property if the property does not
                    // already exist, and you are setting it to Portrait.
                    // That is the default value.
                    if (value != PageOrientationValues.Portrait)
                    {
                        pageOrientationChanged = true;
                        documentChanged = true;
                        pageSize.Orient = new EnumValue<PageOrientationValues>(value);
                    }
                }
                else
                {
                    // The Orient property exists, but its value
                    // is different than the new value.
                    if (pageSize.Orient.Value != value)
                    {
                        pageSize.Orient.Value = value;
                        pageOrientationChanged = true;
                        documentChanged = true;
                    }
                }

                if (pageOrientationChanged)
                {
                    // Changing the orientation is not enough. You must also
                    // change the page size.
                    var width = pageSize.Width;
                    var height = pageSize.Height;
                    pageSize.Width = height;
                    pageSize.Height = width;

                    PageMargin pageMargin = sectionProperty.Descendants<PageMargin>().FirstOrDefault();

                    if (pageMargin != null)
                    {
                        // Rotate margins. Printer settings control how far you
                        // rotate when switching to landscape mode. Not having those
                        // settings, this code rotates 90 degrees. You could easily
                        // modify this behavior, or make it a parameter for the
                        // procedure.
                        var top = pageMargin.Top.Value;
                        var bottom = pageMargin.Bottom.Value;
                        var left = pageMargin.Left.Value;
                        var right = pageMargin.Right.Value;

                        pageMargin.Top = new Int32Value((int)left);
                        pageMargin.Bottom = new Int32Value((int)right);
                        pageMargin.Left = new UInt32Value((uint)System.Math.Max(0, bottom));
                        pageMargin.Right = new UInt32Value((uint)System.Math.Max(0, top));
                    }
                }
            }

            if (documentChanged)
            {
                MainDocumentPart.Document.Save();
            }

            return this;
        }

        internal void Create(WordprocessingDocument doc)
        {
            document = doc;

            doc.AddMainDocumentPart();

            MainDocumentPart.Document = new Document();
            MainDocumentPart.Document.AppendChild(new Body());

            PostCreate();
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

        internal void PostCreate()
        {
            var sectionProperty = Body.GetOrCreate<SectionProperties>();

            var pageSize = sectionProperty.GetOrCreate<PageSize>();
            pageSize.Height = 15840;
            pageSize.Width = 12240;

            // w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"
            var pageMargins = sectionProperty.GetOrCreate<PageMargin>();
            pageMargins.Top = Inch;
            pageMargins.Right = UInch;
            pageMargins.Bottom = Inch;
            pageMargins.Left = UInch;
            pageMargins.Header = UHalfInch;
            pageMargins.Footer = UHalfInch;
            pageMargins.Gutter = UZero;
        }
    }
}
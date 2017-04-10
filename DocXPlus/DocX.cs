using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DocXPlus
{
    public class DocX
    {
        internal const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private WordprocessingDocument document;

        private IList<Models.Footer> footers;
        private IList<Models.Header> headers;

        public PageOrientationValues Orientation
        {
            get
            {
                var sectionProperty = Body.GetOrCreate<SectionProperties>();
                PageSize pageSize = sectionProperty.GetOrCreate<PageSize>();

                return pageSize.Orient ?? PageOrientationValues.Portrait;
            }
            set
            {
                SetOrientation(value);
            }
        }

        public IEnumerable<Models.Paragraph> Paragraphs
        {
            get
            {
                var paragraphs = Body.Descendants<Paragraph>();

                var result = new List<Models.Paragraph>();

                foreach (var paragraph in paragraphs)
                {
                    result.Add(new Models.Paragraph(paragraph));
                }

                return result;
            }
        }

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

        public Models.Footer AddFooter(HeaderFooterValues type)
        { // get the section property for the body
            // which will contain any existing footer references
            var sectionProperty = Body.GetOrCreate<SectionProperties>();

            return AddFooter(type, sectionProperty);
        }

        public Models.Header AddHeader(HeaderFooterValues type)
        {
            // get the section property for the body
            // which will contain any existing header references
            var sectionProperty = Body.GetOrCreate<SectionProperties>();

            return AddHeader(type, sectionProperty);
        }

        public Models.Paragraph AddParagraph()
        {
            var sectionProperties = Body.GetOrCreate<SectionProperties>();
            var paragraph = sectionProperties.InsertBeforeSelf(new Paragraph());

            return new Models.Paragraph(paragraph);
        }

        public void Close()
        {
            Save();

            document.Close();
        }

        public Models.Paragraph InsertPageBreak()
        {
            var paragraph = Body.Descendants<Paragraph>().LastOrDefault();

            if (paragraph == null)
            {
                var sectionProperties = Body.GetOrCreate<SectionProperties>();
                paragraph = sectionProperties.InsertBeforeSelf(new Paragraph());
            }

            paragraph.AppendChild(new Run(new Break() { Type = BreakValues.Page }));

            return new Models.Paragraph(paragraph);
        }

        public void InsertSectionPageBreak()
        {
            // get or create the body section properties
            // we will clone this to create the new section properties
            var bodySectionProperties = Body.GetOrCreate<SectionProperties>();

            // get the last paragraph
            var paragraph = Body.Descendants<Paragraph>().LastOrDefault();

            var addParagraph = paragraph == null;

            if (paragraph != null)
            {
                if (paragraph.Descendants<SectionProperties>().Count() > 0)
                {
                    addParagraph = true;
                }
            }

            if (addParagraph)
            {// no paragraphs or the last paragraph already has a section property
                paragraph = bodySectionProperties.InsertBeforeSelf(new Paragraph());
            }

            // get the paragraph's properties
            var paragraphProperties = paragraph.GetOrCreate<ParagraphProperties>(true);

            // remove title page before the clone
            bodySectionProperties.RemoveAllChildren<TitlePage>();

            // clone the document section properties
            // to get the page size, orientation etc
            var newSectionProperties = (SectionProperties)bodySectionProperties.CloneNode(true);

            // get rid of any header or footer references from
            // the body section properties as they are now in the
            // new section properties
            bodySectionProperties.RemoveAllChildren<HeaderReference>();
            bodySectionProperties.RemoveAllChildren<FooterReference>();

            // add the new section properties to the paragraph properties
            paragraphProperties.AppendChild(newSectionProperties);
        }

        public void Save()
        {
            document.MainDocumentPart.Document.Save();

            if (headers != null)
            {
                foreach (var header in headers)
                {
                    header.Save();
                }
            }

            if (footers != null)
            {
                foreach (var footer in footers)
                {
                    footer.Save();
                }
            }

            document.Save();
        }

        internal static void GenerateFooterPartContent(FooterPart part)
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

        internal static void GenerateHeaderPartContent(HeaderPart part)
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

        internal Models.Footer AddFooter(HeaderFooterValues type, SectionProperties sectionProperty)
        {
            var footerReferences = sectionProperty.Descendants<FooterReference>();

            // footer reference exists for this type?
            var footerReference = footerReferences.FirstOrDefault(p => p.Type == type);

            Models.Footer footer = null;

            if (footerReference == null)
            {
                var part = MainDocumentPart.AddNewPart<FooterPart>();

                var id = MainDocumentPart.GetIdOfPart(part);

                GenerateFooterPartContent(part);

                sectionProperty.RemoveAllChildren<FooterReference>("type", w, type.ToString());

                footerReference = sectionProperty.PrependChild(new FooterReference() { Id = id, Type = type });

                footer = new Models.Footer(part.Footer);
            }

            if (footers == null)
            {
                footers = new List<Models.Footer>();
            }

            footers.Add(footer);

            return footer;
        }

        internal Models.Header AddHeader(HeaderFooterValues type, SectionProperties sectionProperty)
        {
            var headerReferences = sectionProperty.Descendants<HeaderReference>();

            // header reference exists for this type?
            var headerReference = headerReferences.FirstOrDefault(p => p.Type == type);

            Models.Header header = null;

            if (headerReference != null)
            {
                var part = MainDocumentPart.GetPartById(headerReference.Id);

                MainDocumentPart.DeletePart(part);

                sectionProperty.RemoveAllChildren<HeaderReference>("type", w, type.ToString());

                headerReference = null;
            }

            if (headerReference == null)
            {
                var part = MainDocumentPart.AddNewPart<HeaderPart>();

                var id = MainDocumentPart.GetIdOfPart(part);

                GenerateHeaderPartContent(part);

                sectionProperty.RemoveAllChildren<HeaderReference>("type", w, type.ToString());

                headerReference = sectionProperty.PrependChild(new HeaderReference() { Id = id, Type = type });

                header = new Models.Header(part.Header);
            }

            if (headers == null)
            {
                headers = new List<Models.Header>();
            }

            headers.Add(header);

            return header;
        }

        internal void Create(WordprocessingDocument doc)
        {
            document = doc;

            doc.AddMainDocumentPart();

            MainDocumentPart.Document = new Document();
            MainDocumentPart.Document.AppendChild(new Body());

            PostCreate();
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

            var titlePage = sectionProperty.GetOrCreate<TitlePage>();
        }

        internal DocX SetOrientation(PageOrientationValues value)
        {
            bool documentChanged = false;

            var sectionProperty = Body.GetOrCreate<SectionProperties>();

            bool pageOrientationChanged = false;

            PageSize pageSize = sectionProperty.GetOrCreate<PageSize>();

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

            if (documentChanged)
            {
                MainDocumentPart.Document.Save();
            }

            return this;
        }
    }
}
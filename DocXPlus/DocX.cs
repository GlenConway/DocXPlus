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
        internal static MarkupCompatibilityAttributes MarkupCompatibilityAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se wp14" };
        private WordprocessingDocument document;

        private IList<Models.Footer> footers;
        private IList<Models.Header> headers;

        public PageOrientationValues Orientation
        {
            get
            {
                return GetPageSize().Orient ?? PageOrientationValues.Portrait;
            }
            set
            {
                SetOrientation(value);
            }
        }

        public UInt32Value PageHeight
        {
            get
            {
                return GetPageSize().Height;
            }
            set
            {
                GetPageSize().Height = value;
            }
        }

        public UInt32Value PageWidth
        {
            get
            {
                return GetPageSize().Width;
            }
            set
            {
                GetPageSize().Width = value;
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

        Models.PageMargins pageMargins;
        public Models.PageMargins PageMargins { get
            {
                if (pageMargins == null)
                {
                    pageMargins = new Models.PageMargins(this);

                    
                }

                return pageMargins;
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
        {
            // get the section property for the body
            // which will contain any existing footer references
            return AddFooter(type, GetBodySectionProperty());
        }

        public Models.Header AddHeader(HeaderFooterValues type)
        {
            // get the section property for the body
            // which will contain any existing header references
            return AddHeader(type, GetBodySectionProperty());
        }

        /// <summary>
        /// Adds a paragraph to the document just before the body section properties
        /// </summary>
        /// <returns></returns>
        public Models.Paragraph AddParagraph()
        {
            var paragraph = GetBodySectionProperty().InsertBeforeSelf(new Paragraph());

            return new Models.Paragraph(paragraph);
        }

        public Models.Table AddTable(int numberOfColumns)
        {
            var table = GetBodySectionProperty().InsertBeforeSelf(new Table());

            var result = new Models.Table(table, numberOfColumns, this)
            {
                TableStyle = "TableGrid",
                Width = "0",
                WidthType = TableWidthUnitValues.Auto
            };

            result.TableLook.Value = "04A0";
            result.TableLook.FirstRow = true;
            result.TableLook.LastRow = false;
            result.TableLook.FirstColumn = true;
            result.TableLook.LastColumn = false;
            result.TableLook.NoHorizontalBand = false;
            result.TableLook.NoVerticalBand = true;

            return result;
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
                paragraph = GetBodySectionProperty().InsertBeforeSelf(new Paragraph());
            }

            paragraph.AppendChild(new Run(new Break() { Type = BreakValues.Page }));

            return new Models.Paragraph(paragraph);
        }

        public void InsertSectionPageBreak()
        {
            // get or create the body section properties
            // we will clone this to create the new section properties
            var bodySectionProperties = GetBodySectionProperty();

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
            var footer = new Footer() { MCAttributes = MarkupCompatibilityAttributes };

            Schemas.AddNamespaceDeclarations(footer);

            part.Footer = footer;
        }

        internal static void GenerateHeaderPartContent(HeaderPart part)
        {
            var header = new Header() { MCAttributes = MarkupCompatibilityAttributes };

            Schemas.AddNamespaceDeclarations(header);

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

                sectionProperty.RemoveAllChildren<FooterReference>("type", Schemas.w, type.ToString());

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

                sectionProperty.RemoveAllChildren<HeaderReference>("type", Schemas.w, type.ToString());

                headerReference = null;
            }

            if (headerReference == null)
            {
                var part = MainDocumentPart.AddNewPart<HeaderPart>();

                var id = MainDocumentPart.GetIdOfPart(part);

                GenerateHeaderPartContent(part);

                sectionProperty.RemoveAllChildren<HeaderReference>("type", Schemas.w, type.ToString());

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

        internal PageSize GetPageSize()
        {
            return GetBodySectionProperty().GetOrCreate<PageSize>();
        }

        internal void PostCreate()
        {
            var sectionProperty = GetBodySectionProperty();

            // Letter - 8.5" x 11"
            PageHeight = 15840;
            PageWidth = 12240;

            // w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"
            PageMargins.TopAndBottom = Units.Inch;
            PageMargins.RightAndLeft = Units.UInch;

            PageMargins.Header = Units.UHalfInch;
            PageMargins.Footer = Units.UHalfInch;
            PageMargins.Gutter = Units.UZero;

            var titlePage = sectionProperty.GetOrCreate<TitlePage>();
        }

        internal DocX SetOrientation(PageOrientationValues value)
        {
            bool documentChanged = false;

            var sectionProperty = GetBodySectionProperty();

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

        internal SectionProperties GetBodySectionProperty()
        {
            return Body.GetOrCreate<SectionProperties>();
        }
    }
}
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

        private IEnumerable<Footer> footers;
        private IEnumerable<Header> headers;

        private PageMargins pageMargins;
        public Footer DefaultFooter => footers.Where(p => p.Type == HeaderFooterValues.Default).First();
        public Header DefaultHeader => headers.Where(p => p.Type == HeaderFooterValues.Default).First();

        public bool DifferentFirstPage
        {
            get
            {
                if (!GetBodySectionProperty().Has<TitlePage>())
                {
                    return false;
                }

                var titlePage = GetBodySectionProperty().GetOrCreate<TitlePage>();

                return titlePage.Val;
            }
            set
            {
                if (value)
                {
                    var titlePage = GetBodySectionProperty().GetOrCreate<TitlePage>();
                    titlePage.Val = value;
                }
                else
                {
                    GetBodySectionProperty().RemoveAllChildren<TitlePage>();
                }
            }
        }

        public bool EvenAndOddHeaders
        {
            get
            {
                var settings = Settings;

                if (!settings.Has<EvenAndOddHeaders>())
                {
                    return false;
                }

                var evenAndOddHeaders = settings.GetOrCreate<EvenAndOddHeaders>();

                return evenAndOddHeaders.Val;
            }
            set
            {
                var settings = Settings;

                if (value)
                {
                    var evenAndOddHeaders = settings.GetOrCreate<EvenAndOddHeaders>();

                    evenAndOddHeaders.Val = value;
                }
                else
                {
                    settings.RemoveAllChildren<EvenAndOddHeaders>();
                }
            }
        }

        public Footer EvenFooter => footers.Where(p => p.Type == HeaderFooterValues.Even).First();
        public Header EvenHeader => headers.Where(p => p.Type == HeaderFooterValues.Even).First();
        public Footer FirstFooter => footers.Where(p => p.Type == HeaderFooterValues.First).First();
        public Header FirstHeader => headers.Where(p => p.Type == HeaderFooterValues.First).First();

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

        public PageMargins PageMargins
        {
            get
            {
                if (pageMargins == null)
                {
                    pageMargins = new PageMargins(this);
                }

                return pageMargins;
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

        public IEnumerable<Paragraph> Paragraphs
        {
            get
            {
                var paragraphs = Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();

                var result = new List<Paragraph>();

                foreach (var paragraph in paragraphs)
                {
                    result.Add(new Paragraph(paragraph));
                }

                return result;
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

        internal Settings Settings
        {
            get
            {
                var part = document.MainDocumentPart.DocumentSettingsPart;

                if (part == null)
                {
                    part = document.MainDocumentPart.AddNewPart<DocumentSettingsPart>();

                    GenerateDocumentSettingsPartContent(part);

                    part.Settings.Save();
                }

                if (part.Settings == null)
                {
                    part.Settings = new Settings() { MCAttributes = MarkupCompatibilityAttributes };
                    Schemas.AddNamespaceDeclarations(part.Settings);

                    part.Settings.Save();
                }

                return part.Settings;
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

        public void AddFooters()
        {
            var sectionProperty = GetBodySectionProperty();

            var references = sectionProperty.Descendants<FooterReference>();

            foreach (var reference in references)
            {
                var part = MainDocumentPart.GetPartById(reference.Id);

                MainDocumentPart.DeletePart(reference.Id);
            }

            sectionProperty.RemoveAllChildren<FooterReference>();

            footers = new List<Footer>
            {
                AddFooter(HeaderFooterValues.Default),
                AddFooter(HeaderFooterValues.Even),
                AddFooter(HeaderFooterValues.First)
            };
        }

        public void AddHeaders()
        {
            var sectionProperty = GetBodySectionProperty();

            var references = sectionProperty.Descendants<HeaderReference>();

            foreach (var reference in references)
            {
                var part = MainDocumentPart.GetPartById(reference.Id);

                MainDocumentPart.DeletePart(reference.Id);
            }

            sectionProperty.RemoveAllChildren<HeaderReference>();

            headers = new List<Header>
            {
                AddHeader(HeaderFooterValues.Default),
                AddHeader(HeaderFooterValues.Even),
                AddHeader(HeaderFooterValues.First)
            };
        }

        /// <summary>
        /// Adds a paragraph to the document just before the body section properties
        /// </summary>
        /// <returns></returns>
        public Paragraph AddParagraph()
        {
            var paragraph = GetBodySectionProperty().InsertBeforeSelf(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());

            return new Paragraph(paragraph);
        }

        /// <summary>
        /// Adds a Table to the document with the specified number of columns
        /// </summary>
        /// <param name="numberOfColumns"></param>
        /// <returns></returns>
        public Table AddTable(int numberOfColumns)
        {
            var table = GetBodySectionProperty().InsertBeforeSelf(new DocumentFormat.OpenXml.Wordprocessing.Table());

            return AddTable(numberOfColumns, table);
        }

        public void Close()
        {
            Save();

            document.Close();
        }

        public Paragraph InsertPageBreak()
        {
            var paragraph = Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().LastOrDefault();

            if (paragraph == null)
            {
                paragraph = GetBodySectionProperty().InsertBeforeSelf(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
            }

            paragraph.AppendChild(new Run(new Break() { Type = BreakValues.Page }));

            return new Paragraph(paragraph);
        }

        public void InsertSectionPageBreak()
        {
            // first save any header or footer content because after
            // this, there will be a new set of headers and footers
            SaveHeaders();
            SaveFooters();

            // get or create the body section properties
            // we will clone this to create the new section properties
            var bodySectionProperties = GetBodySectionProperty();

            // get the last paragraph
            var paragraph = Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().LastOrDefault();

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
                paragraph = bodySectionProperties.InsertBeforeSelf(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
            }

            // get the paragraph's properties
            var paragraphProperties = paragraph.GetOrCreate<ParagraphProperties>(true);

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
            Settings.Save();

            MainDocumentPart.Document.Save();

            SaveHeaders();
            SaveFooters();

            document.Save();
        }

        internal static void GenerateFooterPartContent(FooterPart part)
        {
            var footer = new DocumentFormat.OpenXml.Wordprocessing.Footer() { MCAttributes = MarkupCompatibilityAttributes };

            Schemas.AddNamespaceDeclarations(footer);

            part.Footer = footer;
        }

        internal static void GenerateHeaderPartContent(HeaderPart part)
        {
            var header = new DocumentFormat.OpenXml.Wordprocessing.Header() { MCAttributes = MarkupCompatibilityAttributes };

            Schemas.AddNamespaceDeclarations(header);

            part.Header = header;
        }

        internal Table AddTable(int numberOfColumns, DocumentFormat.OpenXml.Wordprocessing.Table table)
        {
            var result = new Table(table, numberOfColumns, this)
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

        internal void Create(WordprocessingDocument doc)
        {
            document = doc;

            doc.AddMainDocumentPart();

            MainDocumentPart.Document = new Document();
            MainDocumentPart.Document.AppendChild(new Body());

            PostCreate();
        }

        internal SectionProperties GetBodySectionProperty()
        {
            return Body.GetOrCreate<SectionProperties>();
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

            var documentSettingsPart = MainDocumentPart.AddNewPart<DocumentSettingsPart>();
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

        private static void GenerateDocumentSettingsPartContent(DocumentSettingsPart documentSettingsPart)
        {
            Settings settings = new Settings() { MCAttributes = MarkupCompatibilityAttributes };
            Schemas.AddNamespaceDeclarations(settings);

            documentSettingsPart.Settings = settings;
        }

        private Footer AddFooter(HeaderFooterValues type)
        {
            var part = MainDocumentPart.AddNewPart<FooterPart>();

            var id = MainDocumentPart.GetIdOfPart(part);

            GenerateFooterPartContent(part);

            GetBodySectionProperty().PrependChild(new FooterReference() { Id = id, Type = type });

            return new Footer(part.Footer, this, type);
        }

        private Header AddHeader(HeaderFooterValues type)
        {
            var part = MainDocumentPart.AddNewPart<HeaderPart>();

            var id = MainDocumentPart.GetIdOfPart(part);

            GenerateHeaderPartContent(part);

            GetBodySectionProperty().PrependChild(new HeaderReference() { Id = id, Type = type });

            return new Header(part.Header, this, type);
        }

        private void SaveFooters()
        {
            if (footers == null)
            {
                return;
            }

            foreach (var footer in footers)
            {
                footer.Save();
            }
        }

        private void SaveHeaders()
        {
            if (headers == null)
            {
                return;
            }

            foreach (var header in headers)
            {
                header.Save();
            }
        }
    }
}
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocXPlus
{
    /// <summary>
    /// Wrapper around an OpenXml SDK Wordprocessing Document
    /// </summary>
    public class DocX : IDisposable
    {
        internal static MarkupCompatibilityAttributes MarkupCompatibilityAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se wp14" };

        private bool disposed = false;

        private WordprocessingDocument document;

        private IEnumerable<Footer> footers;

        private IEnumerable<Header> headers;

        private PageMargins pageMargins;

        private Stream stream;

        /// <summary>
        /// Default destructor
        /// </summary>
        ~DocX()
        {
            Dispose(false);
        }

        /// <summary>
        /// The default (odd) footer.
        /// </summary>
        public Footer DefaultFooter => footers.Where(p => p.Type == HeaderFooterValues.Default).First();

        /// <summary>
        /// The default (odd) header.
        /// </summary>
        public Header DefaultHeader => headers.Where(p => p.Type == HeaderFooterValues.Default).First();

        /// <summary>
        /// Specify whether the first page has a different header than the rest of the document
        /// </summary>
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

        /// <summary>
        /// Specify if even and odd pages use different headers / footers.
        /// </summary>
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

        /// <summary>
        /// Even page footer
        /// </summary>
        public Footer EvenFooter => footers.Where(p => p.Type == HeaderFooterValues.Even).First();

        /// <summary>
        /// Even page header
        /// </summary>
        public Header EvenHeader => headers.Where(p => p.Type == HeaderFooterValues.Even).First();

        /// <summary>
        /// First page footer
        /// </summary>
        public Footer FirstFooter => footers.Where(p => p.Type == HeaderFooterValues.First).First();

        /// <summary>
        /// First page header
        /// </summary>
        public Header FirstHeader => headers.Where(p => p.Type == HeaderFooterValues.First).First();

        /// <summary>
        /// Orientation of the document or current section
        /// </summary>
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

        /// <summary>
        /// Height of the page in Twips
        /// </summary>
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

        /// <summary>
        /// Specifies the page margins
        /// </summary>
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

        /// <summary>
        /// Width of the page in Twips
        /// </summary>
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

        /// <summary>
        /// All of the paragraphs in the document
        /// </summary>
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
                if (document == null)
                {
                    throw new InvalidOperationException("You must call Create before accessing the Main Document Part.");
                }

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

        /// <summary>
        /// Creates a new document using the supplied path and type
        /// </summary>
        /// <param name="path"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static DocX Create(string path, WordprocessingDocumentType type)
        {
            return Create(path, type, false);
        }

        /// <summary>
        /// Creates a new document using the supplied path, type and autosave value
        /// </summary>
        /// <param name="path"></param>
        /// <param name="type"></param>
        /// <param name="autoSave"></param>
        /// <returns></returns>
        public static DocX Create(string path, WordprocessingDocumentType type, bool autoSave)
        {
            var docX = new DocX();
            docX.Create(WordprocessingDocument.Create(path, type, autoSave));

            return docX;
        }

        /// <summary>
        /// Creates a new document using the supplied stream and type
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static DocX Create(Stream stream, WordprocessingDocumentType type)
        {
            return Create(stream, type, false);
        }

        /// <summary>
        /// Creates a new document using the supplied stream, type and autosave value
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="type"></param>
        /// <param name="autoSave"></param>
        /// <returns></returns>
        public static DocX Create(Stream stream, WordprocessingDocumentType type, bool autoSave)
        {
            var docX = new DocX();
            docX.Create(WordprocessingDocument.Create(stream, type, autoSave));

            return docX;
        }

        /// <summary>
        /// Adds three footers to the document (default, even and first).
        /// </summary>
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

        /// <summary>
        /// Adds three headers to the document (default, even and first).
        /// </summary>
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
        /// Adds an image to the drawing which can then be added to a paragraph
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="x">The width of the image in English Metric Units (EMU)</param>
        /// <param name="y">The height of the image in English Metric Units (EMU)</param>
        /// <returns></returns>
        public Drawing AddImage(string fileName, Int64Value x, Int64Value y)
        {
            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                return AddImage(stream, FileNameContentType(fileName), x, y);
            }
        }

        /// <summary>
        /// Adds an image to the drawing which can then be added to a paragraph
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="contentType"></param>
        /// <param name="width">The width of the image in English Metric Units (EMU)</param>
        /// <param name="height">The height of the image in English Metric Units (EMU)</param>
        /// <returns></returns>
        public Drawing AddImage(Stream stream, string contentType, Int64Value width, Int64Value height)
        {
            ImagePart imagePart = MainDocumentPart.AddImagePart(contentType);
            imagePart.FeedData(stream);

            return CreateDrawing(MainDocumentPart.GetIdOfPart(imagePart), width, height);
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
        /// Adds a paragraph with the supplied text to the document.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public Paragraph AddParagraph(string text)
        {
            return AddParagraph().Append(text);
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

        /// <summary>
        /// Adds a Table to the document with the specified number of columns using the percent widths
        /// </summary>
        /// <param name="numberOfColumns"></param>
        /// <param name="percent"></param>
        /// <returns></returns>
        public Table AddTable(int numberOfColumns, params int[] percent)
        {
            var table = GetBodySectionProperty().InsertBeforeSelf(new DocumentFormat.OpenXml.Wordprocessing.Table());

            return AddTable(numberOfColumns, table, percent);
        }

        /// <summary>
        /// Adds a Table to the document with the specified number of columns using the supplied widths
        /// </summary>
        /// <param name="numberOfColumns"></param>
        /// <param name="widths">The widths of the columns in Twips</param>
        /// <returns></returns>
        public Table AddTable(int numberOfColumns, params string[] widths)
        {
            var table = GetBodySectionProperty().InsertBeforeSelf(new DocumentFormat.OpenXml.Wordprocessing.Table());

            return AddTable(numberOfColumns, table, widths);
        }

        /// <summary>
        /// Saves and closes the document.
        /// </summary>
        public void Close()
        {
            Save();

            document.Close();
        }

        /// <summary>
        /// Creates a new document of the supplied type using a stream
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public DocX Create(WordprocessingDocumentType type)
        {
            stream = new MemoryStream();

            return Create(stream, type, false);
        }

        /// <summary>
        /// Creates a new document of type Document using a stream
        /// </summary>
        /// <returns></returns>
        public void Create()
        {
            stream = new MemoryStream();

            Create(WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, false));
        }

        /// <summary>
        /// Disposes the document
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Inserts a Page Break
        /// </summary>
        /// <returns></returns>
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

        /// <summary>
        /// Inserts a Section Page Break
        /// </summary>
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
            {
                // no paragraphs or the last paragraph already has a section property
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

        /// <summary>
        /// Saves the document to the underlying stream. Does not write out the document to the file system until Close() is called.
        /// </summary>
        public void Save()
        {
            Settings.Save();

            MainDocumentPart.Document.Save();

            SaveHeaders();
            SaveFooters();

            document.Save();
        }

        /// <summary>
        /// Saves the document to the supplied stream
        /// </summary>
        /// <param name="stream"></param>
        public void SaveAs(Stream stream)
        {
            Save();

            document.Clone(stream);
        }

        internal static string FileNameContentType(string fileName)
        {
            switch (Path.GetExtension(fileName))
            {
                case ".tif":
                case ".tiff":
                    return "image/tif";

                case ".bmp":
                case ".png":
                    return "image/png";

                case ".gif":
                    return "image/gif";

                case ".jpeg":
                    return "image/jpeg";
            }

            return "image/jpg";
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

            SetTableLook(result);

            return result;
        }

        internal Table AddTable(int numberOfColumns, DocumentFormat.OpenXml.Wordprocessing.Table table, params int[] percent)
        {
            var result = new Table(table, numberOfColumns, this, percent)
            {
                TableStyle = "TableGrid",
                Width = "0",
                WidthType = TableWidthUnitValues.Auto
            };

            SetTableLook(result);

            return result;
        }

        internal Table AddTable(int numberOfColumns, DocumentFormat.OpenXml.Wordprocessing.Table table, params string[] widths)
        {
            var result = new Table(table, numberOfColumns, this, widths)
            {
                TableStyle = "TableGrid",
                Width = "0",
                WidthType = TableWidthUnitValues.Auto
            };

            SetTableLook(result);

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

        /// <summary>
        /// Disposing
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    // Dispose managed resources.
                    if (stream != null)
                    {
                        stream.Flush();
                        stream.Dispose();

                        stream = null;
                    }
                }

                // There are no unmanaged resources to release, but
                // if we add them, they need to be released here.
            }

            disposed = true;
        }

        private static void GenerateDocumentSettingsPartContent(DocumentSettingsPart documentSettingsPart)
        {
            Settings settings = new Settings() { MCAttributes = MarkupCompatibilityAttributes };
            Schemas.AddNamespaceDeclarations(settings);

            documentSettingsPart.Settings = settings;
        }

        private static void SetTableLook(Table result)
        {
            result.TableLook.Value = "04A0";
            result.TableLook.FirstRow = true;
            result.TableLook.LastRow = false;
            result.TableLook.FirstColumn = true;
            result.TableLook.LastColumn = false;
            result.TableLook.NoHorizontalBand = false;
            result.TableLook.NoVerticalBand = true;
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

        private Drawing CreateDrawing(string id, Int64Value width, Int64Value height)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = width, Cy = height },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = 1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = 0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = id,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = width, Cy = height }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            return element;
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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace DocXPlus
{
    /// <summary>
    /// Container class for common functions
    /// </summary>
    public abstract class Container
    {
        /// <summary>
        /// The width of the container
        /// </summary>
        public abstract UInt32Value AvailableWidth
        {
            get;
        }

        /// <summary>
        /// Adds an image to the container which can then be added to a paragraph
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
        /// Adds an image to the container which can then be added to a paragraph
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="contentType"></param>
        /// <param name="width">The width of the image in English Metric Units (EMU)</param>
        /// <param name="height">The height of the image in English Metric Units (EMU)</param>
        /// <returns></returns>
        public Drawing AddImage(Stream stream, string contentType, Int64Value width, Int64Value height)
        {
            return DocX.CreateDrawing(AddImagePart(stream, contentType), width, height);
        }

        /// <summary>
        /// Adds an image to the container which can then be added to a paragraph
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
        /// Adds a paragraph
        /// </summary>
        /// <returns></returns>
        public Paragraph AddParagraph()
        {
            return new Paragraph(NewParagraph());
        }

        /// <summary>
        /// Adds a paragraph with the supplied text
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public Paragraph AddParagraph(string text)
        {
            return AddParagraph().Append(text);
        }

        /// <summary>
        /// Adds a paragraph with the supplied text and sets the alignment
        /// </summary>
        /// <param name="text"></param>
        /// <param name="alignment"></param>
        /// <returns></returns>
        public Paragraph AddParagraph(string text, Align alignment)
        {
            return AddParagraph().Append(text).SetAlignment(alignment);
        }

        /// <summary>
        /// Adds a Table to the container with the specified number of columns
        /// </summary>
        /// <param name="numberOfColumns"></param>
        /// <returns></returns>
        public Table AddTable(int numberOfColumns)
        {
            return AddTable(numberOfColumns, NewTable());
        }

        /// <summary>
        /// Adds a Table to the container with the specified number of columns using the percent widths
        /// </summary>
        /// <param name="numberOfColumns"></param>
        /// <param name="percent"></param>
        /// <returns></returns>
        public Table AddTable(int numberOfColumns, params int[] percent)
        {
            return AddTable(numberOfColumns, NewTable(), percent);
        }

        /// <summary>
        /// Adds a Table to the container with the specified number of columns using the supplied widths
        /// </summary>
        /// <param name="numberOfColumns"></param>
        /// <param name="widths">The widths of the columns in Twips, CM or Inch. Leave blank for Twips, end with cm for centimetres or in for inches</param>
        /// <returns></returns>
        public Table AddTable(int numberOfColumns, params string[] widths)
        {
            return AddTable(numberOfColumns, NewTable(), widths);
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
            for (int i = 0; i < widths.Length; i++)
            {
                var width = widths[i];

                if (width.EndsWith("cm", System.StringComparison.OrdinalIgnoreCase))
                {
                    width = width.Remove(width.Length - 2, 2);
                    
                    if (double.TryParse(width, out double value))
                    {
                        width = Units.CMToTwips(value).Value.ToString();
                    }
                }

                if (width.EndsWith("in", System.StringComparison.OrdinalIgnoreCase))
                {
                    width = width.Remove(width.Length - 2, 2);

                    if (double.TryParse(width, out double value))
                    {
                        width = Units.InchToTwips(value).Value.ToString();
                    }
                }
            }

            var result = new Table(table, numberOfColumns, this, widths)
            {
                TableStyle = "TableGrid",
                Width = "0",
                WidthType = TableWidthUnitValues.Auto
            };

            SetTableLook(result);

            return result;
        }

        /// <summary>
        /// Adds an image part to the container and returns the part ID
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="contentType"></param>
        /// <returns></returns>
        protected abstract string AddImagePart(Stream stream, string contentType);

        /// <summary>
        /// Creates a new paragraph in the container
        /// </summary>
        /// <returns></returns>
        protected abstract DocumentFormat.OpenXml.Wordprocessing.Paragraph NewParagraph();

        /// <summary>
        /// Creates a new table in the container
        /// </summary>
        /// <returns></returns>
        protected abstract DocumentFormat.OpenXml.Wordprocessing.Table NewTable();

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
    }
}
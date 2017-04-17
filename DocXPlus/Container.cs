using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

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
        /// <param name="widths">The widths of the columns in Twips</param>
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
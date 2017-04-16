namespace DocXPlus
{
    /// <summary>
    /// Helper extensions when working with tables
    /// </summary>
    public static class TableExtensions
    {
        /// <summary>
        /// Double underlines each cell in a row from startIndex to endIndex
        /// </summary>
        /// <param name="row"></param>
        /// <param name="startIndex"></param>
        /// <param name="endIndex"></param>
        public static void DoubleUnderline(this TableRow row, int startIndex, int endIndex)
        {
            for (int i = startIndex; i <= endIndex; i++)
            {
                row.Cells[i].DoubleUnderline();
            }
        }

        /// <summary>
        /// Double underlines a cell
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static TableCell DoubleUnderline(this TableCell cell)
        {
            cell.Borders.BottomBorder.Set(Units.OnePt, BorderValue.Double);

            return cell;
        }

        /// <summary>
        /// Overlines a cell
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static TableCell Overline(this TableCell cell)
        {
            cell.Borders.TopBorder.Set(Units.OnePt, BorderValue.Single);

            return cell;
        }

        /// <summary>
        /// Sets the text of the first paragraph in the row cell at the given index
        /// </summary>
        /// <param name="row"></param>
        /// <param name="cellIndex"></param>
        /// <param name="text"></param>
        /// <returns></returns>
        public static Paragraph SetBoldText(this TableRow row, int cellIndex, string text)
        {
            return row.Cells[cellIndex].SetBoldText(text);
        }

        /// <summary>
        /// Sets the text of the first paragraph in the row cell at the given index
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="text"></param>
        /// <returns></returns>
        public static Paragraph SetBoldText(this TableCell cell, string text)
        {
            return cell.SetText(text, true);
        }

        /// <summary>
        /// Sets the text of the first paragraph in the row cell at the given index
        /// </summary>
        /// <param name="row"></param>
        /// <param name="cellIndex"></param>
        /// <param name="text"></param>
        /// <param name="alignment"></param>
        /// <returns></returns>
        public static Paragraph SetBoldText(this TableRow row, int cellIndex, string text, Align alignment)
        {
            return row.Cells[cellIndex].AddParagraph(text).Bold().SetAlignment(alignment);
        }

        /// <summary>
        /// Sets the text of the first paragraph in the row cell at the given index
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="text"></param>
        /// <param name="alignment"></param>
        /// <returns></returns>
        public static Paragraph SetBoldText(this TableCell cell, string text, Align alignment)
        {
            return cell.SetText(text, alignment, true);
        }

        /// <summary>
        /// Sets the text of the first paragraph in the row cell at the given index
        /// </summary>
        /// <param name="row"></param>
        /// <param name="cellIndex"></param>
        /// <param name="text"></param>
        /// <returns></returns>
        public static Paragraph SetText(this TableRow row, int cellIndex, string text)
        {
            return row.Cells[cellIndex].SetText(text);
        }

        /// <summary>
        /// Sets the text of the first paragraph in the row cell at the given index
        /// </summary>
        /// <param name="row"></param>
        /// <param name="cellIndex"></param>
        /// <param name="text"></param>
        /// <param name="alignment"></param>
        /// <returns></returns>
        public static Paragraph SetText(this TableRow row, int cellIndex, string text, Align alignment)
        {
            return row.Cells[cellIndex].SetText(text, alignment);
        }

        /// <summary>
        /// Sets the text of the first paragraph in a cell
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="text"></param>
        /// <returns></returns>
        public static Paragraph SetText(this TableCell cell, string text)
        {
            return cell.SetText(text, false);
        }

        /// <summary>
        /// Sets the text of the first paragraph in a cell
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="text"></param>
        /// <param name="bold"></param>
        /// <returns></returns>
        public static Paragraph SetText(this TableCell cell, string text, bool bold)
        {
            return cell.SetText(text, Align.Left, bold);
        }

        /// <summary>
        /// Sets the text of the first paragraph in a cell
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="text"></param>
        /// <param name="alignment"></param>
        /// <returns></returns>
        public static Paragraph SetText(this TableCell cell, string text, Align alignment)
        {
            return cell.SetText(text, alignment, false);
        }

        /// <summary>
        /// Sets the text of the first paragraph in a cell
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="text"></param>
        /// <param name="alignment"></param>
        /// <param name="bold"></param>
        /// <returns></returns>
        public static Paragraph SetText(this TableCell cell, string text, Align alignment, bool bold)
        {
            var paragraph = cell.Paragraphs[0].Append(text);

            if (bold)
            {
                paragraph.Bold();
            }

            paragraph.Alignment = alignment;

            return paragraph;
        }

        /// <summary>
        /// Underlines a cell in a row at the supplied index
        /// </summary>
        /// <param name="row"></param>
        /// <param name="index"></param>
        public static void Underline(this TableRow row, int index)
        {
            row.Underline(index, index);
        }

        /// <summary>
        /// Underlines cells in a row from startIndex to endIndex
        /// </summary>
        /// <param name="row"></param>
        /// <param name="startIndex"></param>
        /// <param name="endIndex"></param>
        public static void Underline(this TableRow row, int startIndex, int endIndex)
        {
            for (int i = startIndex; i <= endIndex; i++)
            {
                row.Cells[i].Underline();
            }
        }

        /// <summary>
        /// Underlines a table cell
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static TableCell Underline(this TableCell cell)
        {
            cell.Borders.BottomBorder.Set(Units.OnePt, BorderValue.Single);

            return cell;
        }

        /// <summary>
        /// Underlines and overlines a cell in a row at the supplied index
        /// </summary>
        /// <param name="row"></param>
        /// <param name="index"></param>
        public static void UnderlineOverline(this TableRow row, int index)
        {
            row.UnderlineOverline(index, index);
        }

        /// <summary>
        /// Underlines and overlines table cells from startIndex to endIndex
        /// </summary>
        /// <param name="row"></param>
        /// <param name="startIndex"></param>
        /// <param name="endIndex"></param>
        public static void UnderlineOverline(this TableRow row, int startIndex, int endIndex)
        {
            for (int i = startIndex; i <= endIndex; i++)
            {
                row.Cells[i].UnderlineOverline();
            }
        }

        /// <summary>
        /// Underlines and overlines a table cell
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static TableCell UnderlineOverline(this TableCell cell)
        {
            cell.Underline();
            return cell.Overline();
        }
    }
}
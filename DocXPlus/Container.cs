namespace DocXPlus
{
    /// <summary>
    /// Container class for common functions
    /// </summary>
    public abstract class Container
    {
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
        /// Creates a new paragraph in the container
        /// </summary>
        /// <returns></returns>
        protected abstract DocumentFormat.OpenXml.Wordprocessing.Paragraph NewParagraph();
    }
}
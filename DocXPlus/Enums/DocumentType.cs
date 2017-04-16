namespace DocXPlus
{
    /// <summary>
    /// Defines the type of word processing document
    /// </summary>
    public enum DocumentType
    {
        /// <summary>
        /// Word Document (*.docx).
        /// </summary>
        Document = 0,

        /// <summary>
        /// Word Template (*.dotx).
        /// </summary>
        Template = 1,

        /// <summary>
        /// Word Macro-Enabled Document (*.docm).
        /// </summary>
        MacroEnabledDocument = 2,

        /// <summary>
        ///  Word Macro-Enabled Template (*.dotm).
        /// </summary>
        MacroEnabledTemplate = 3
    }
}
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

namespace DocXPlus
{
    /// <summary>
    /// Represents the styles for the document
    /// </summary>
    public class Styles
    {
        private DocumentFormat.OpenXml.Wordprocessing.Styles styles;

        internal Styles(DocumentFormat.OpenXml.Wordprocessing.Styles styles)
        {
            this.styles = styles;
        }

        /// <summary>
        /// Normal style
        /// </summary>
        public Style Normal
        {
            get
            {
                return DocumentStyle("Normal");
            }
        }
        
        /// <summary>
        /// Returns the document style specified by the styleId
        /// </summary>
        /// <param name="styleId"></param>
        /// <returns></returns>
        public Style DocumentStyle(string styleId)
        {
            var element = styles.Elements<DocumentFormat.OpenXml.Wordprocessing.Style>().Where(p => p.StyleId == styleId).FirstOrDefault();

            if (element == null)
            {
                element = new DocumentFormat.OpenXml.Wordprocessing.Style()
                {
                    StyleId = styleId,
                    Type = StyleValues.Paragraph,
                    StyleName = new StyleName() { Val = styleId }
                };

                styles.AppendChild(element);
            }

            return new Style(element);
        }

        internal static void AddStylesDefault(DocumentFormat.OpenXml.Wordprocessing.Styles styles)
        {
            var docDefaults = styles.GetOrCreate<DocDefaults>(true);

            var runPropertiesDefault = docDefaults.GetOrCreate<RunPropertiesDefault>(true);
            var runPropertiesBaseStyle = runPropertiesDefault.GetOrCreate<RunPropertiesBaseStyle>();

            var runFonts = runPropertiesBaseStyle.GetOrCreate<RunFonts>();
            runFonts.AsciiTheme = ThemeFontValues.MinorHighAnsi;
            runFonts.EastAsiaTheme = ThemeFontValues.MinorHighAnsi;
            runFonts.ComplexScriptTheme = ThemeFontValues.MinorBidi;

            runPropertiesBaseStyle.GetOrCreate<FontSize>().Val = "22";
            runPropertiesBaseStyle.GetOrCreate<FontSizeComplexScript>().Val = "22";

            var lang = runPropertiesBaseStyle.GetOrCreate<Languages>();
            lang.Val = "en-US";
            lang.EastAsia = "en-US";
            lang.Bidi = "ar-SA";

            var paragraphPropertiesDefault = docDefaults.GetOrCreate<ParagraphPropertiesDefault>();
            var paragraphProperties = paragraphPropertiesDefault.GetOrCreate<ParagraphPropertiesBaseStyle>();
            var spacing = paragraphProperties.GetOrCreate<SpacingBetweenLines>();

            spacing.After = "160";
            spacing.Line = "259";
            spacing.LineRule = LineSpacingRuleValues.Auto;
        }

        internal void CreateStandardStyles()
        {
            var style = DocumentStyle("Normal");
            style.Name = "Normal";
            style.Default = true;
            style.Type = StyleValue.Paragraph;

            style = DocumentStyle("DefaultParagraphFont");
            style.Name = "Default Paragraph Font";
            style.Type = StyleValue.Character;
            style.UIPriority = 1;
            style.Default = true;

            style = DocumentStyle("TableNormal");
            style.Name = "Normal Table";
            style.Type = StyleValue.Character;
            style.UIPriority = 99;
            style.Default = true;
        }

        internal void Save()
        {
            styles.Save();
        }
    }
}
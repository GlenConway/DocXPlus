using DocumentFormat.OpenXml.Wordprocessing;

namespace DocXPlus
{
    /// <summary>
    /// Represents a style in the document
    /// </summary>
    public class Style
    {
        private DocumentFormat.OpenXml.Wordprocessing.Style style;

        internal Style(DocumentFormat.OpenXml.Wordprocessing.Style style)
        {
            this.style = style;
        }

        /// <summary>
        /// Default Style.
        /// </summary>
        public bool Default
        {
            get
            {
                return style.Default;
            }
            set
            {
                style.Default = value;
            }
        }

        /// <summary>
        /// Primary Style Name.
        /// </summary>
        public string Name
        {
            get
            {
                var styleName = style.GetOrCreate<StyleName>(true);
                return styleName.Val;
            }
            set
            {
                var styleName = style.GetOrCreate<StyleName>(true);
                styleName.Val = value;
            }
        }

        /// <summary>
        /// Gets or sets the font size in points
        /// </summary>
        public int Size
        {
            get
            {
                var value = GetFontSize().Val.Value;

                if (int.TryParse(value, out int size))
                {
                    return size * 2;
                }

                return 0;
            }
            set
            {
                GetFontSize().Val = (value * 2).ToString();
            }
        }

        /// <summary>
        /// Style ID.
        /// </summary>
        public string StyleId
        {
            get
            {
                return style.StyleId;
            }
            set
            {
                style.StyleId = value;
            }
        }

        /// <summary>
        ///
        /// </summary>
        public TableCellMarginDefault TableDefaultMargins
        {
            get
            {
                return new TableCellMarginDefault(GetTableProperties().GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.TableCellMarginDefault>());
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public TableIndentation TableIndentation
        {
            get
            {
                return new TableIndentation(GetTableProperties().GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.TableIndentation>());
            }
        }
        /// <summary>
        /// Style Type.
        /// </summary>
        public StyleValue Type
        {
            get
            {
                return Convert.ToStyleValue(style.Type);
            }
            set
            {
                style.Type = Convert.ToStyleValues(value);
            }
        }

        /// <summary>
        /// Optional User Interface Sorting Order.
        /// </summary>
        public int UIPriority
        {
            get
            {
                var uiPriority = style.GetOrCreate<UIPriority>();

                return uiPriority.Val;
            }
            set
            {
                var uiPriority = style.GetOrCreate<UIPriority>();
                uiPriority.Val = value;
            }
        }

        private FontSize GetFontSize()
        {
            return GetRunProperties().GetOrCreate<FontSize>();
        }

        private RunProperties GetRunProperties()
        {
            return style.GetOrCreate<RunProperties>();
        }

        private StyleTableProperties GetTableProperties()
        {
            return style.GetOrCreate<StyleTableProperties>();
        }

        ///// <summary>
        ///// Style Paragraph Properties.
        ///// </summary>
        //public string StyleParagraphProperties
        //{
        //    get
        //    {
        //        return style.StyleParagraphProperties;
        //    }
        //    set
        //    {
        //        style.StyleParagraphProperties
        //        style.StyleParagraphProperties = value;
        //    }
        //}
    }
}
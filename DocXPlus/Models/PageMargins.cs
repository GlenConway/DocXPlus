using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocXPlus
{
    /// <summary>
    /// Represents the page margins
    /// </summary>
    public class PageMargins
    {
        private DocX document;
        private PageMargin pageMargin;

        internal PageMargins(DocX document)
        {
            this.document = document;
            pageMargin = document.GetBodySectionProperty().GetOrCreate<PageMargin>();
        }

        /// <summary>
        /// Bottom margin in Twips
        /// </summary>
        public Int32Value Bottom
        {
            get
            {
                return pageMargin.Bottom;
            }
            set
            {
                pageMargin.Bottom = value;
            }
        }

        /// <summary>
        /// Footer margin in Twips
        /// </summary>
        public UInt32Value Footer
        {
            get
            {
                return pageMargin.Footer;
            }
            set
            {
                pageMargin.Footer = value;
            }
        }

        /// <summary>
        /// Gutter margin in Twips
        /// </summary>
        public UInt32Value Gutter
        {
            get
            {
                return pageMargin.Gutter;
            }
            set
            {
                pageMargin.Gutter = value;
            }
        }

        /// <summary>
        /// Header margin in Twips
        /// </summary>
        public UInt32Value Header
        {
            get
            {
                return pageMargin.Header;
            }
            set
            {
                pageMargin.Header = value;
            }
        }

        /// <summary>
        /// Left margin in Twips
        /// </summary>
        public UInt32Value Left
        {
            get
            {
                return pageMargin.Left;
            }
            set
            {
                pageMargin.Left = value;
            }
        }

        /// <summary>
        /// Right margin in Twips
        /// </summary>
        public UInt32Value Right
        {
            get
            {
                return pageMargin.Right;
            }
            set
            {
                pageMargin.Right = value;
            }
        }

        /// <summary>
        /// Right and left margin in Twips
        /// </summary>
        public UInt32Value RightAndLeft
        {
            get
            {
                return Right.Value + Left.Value;
            }
            set
            {
                Right = value;
                Left = value;
            }
        }

        /// <summary>
        /// Top margin in Twips
        /// </summary>
        public Int32Value Top
        {
            get
            {
                return pageMargin.Top;
            }
            set
            {
                pageMargin.Top = value;
            }
        }

        /// <summary>
        /// Top and bottom margin in Twips
        /// </summary>
        public Int32Value TopAndBottom
        {
            get
            {
                return Top.Value + Bottom.Value;
            }
            set
            {
                Top = value;
                Bottom = value;
            }
        }
    }
}
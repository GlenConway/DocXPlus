using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocXPlus
{
    public class PageMargins
    {
        private DocX document;
        private PageMargin pageMargin;

        public PageMargins(DocX document)
        {
            this.document = document;
            pageMargin = document.GetBodySectionProperty().GetOrCreate<PageMargin>();
        }

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
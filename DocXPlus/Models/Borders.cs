using DocumentFormat.OpenXml;

namespace DocXPlus.Models
{
    public class Borders
    {
        private OpenXmlCompositeElement parent;

        internal Borders(OpenXmlCompositeElement parent)
        {
            this.parent = parent;
        }

        public Border BottomBorder
        {
            get
            {
                var item = parent.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.BottomBorder>();
                return new Border(item);
            }
        }

        public Border EndBorder
        {
            get
            {
                var item = parent.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.EndBorder>();
                return new Border(item);
            }
        }

        public Border InsideHorizontalBorder
        {
            get
            {
                var item = parent.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder>();
                return new Border(item);
            }
        }

        public Border InsideVerticalBorder
        {
            get
            {
                var item = parent.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder>();
                return new Border(item);
            }
        }

        public Border LeftBorder
        {
            get
            {
                var item = parent.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.LeftBorder>();
                return new Border(item);
            }
        }

        public Border RightBorder
        {
            get
            {
                var item = parent.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.RightBorder>();
                return new Border(item);
            }
        }

        public Border StartBorder
        {
            get
            {
                var item = parent.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.StartBorder>();
                return new Border(item);
            }
        }

        public Border TopBorder
        {
            get
            {
                var item = parent.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.TopBorder>();
                return new Border(item);
            }
        }

        public Border TopLeftToBottomRightCellBorder
        {
            get
            {
                var item = parent.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.TopLeftToBottomRightCellBorder>();
                return new Border(item);
            }
        }

        public Border TopRightToBottomLeftCellBorder
        {
            get
            {
                var item = parent.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.TopRightToBottomLeftCellBorder>();
                return new Border(item);
            }
        }
    }
}
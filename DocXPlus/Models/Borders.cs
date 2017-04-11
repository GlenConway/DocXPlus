using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

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
                var item = parent.GetOrCreate<BottomBorder>();
                return new Border(item);
            }
        }

        public Border EndBorder
        {
            get
            {
                var item = parent.GetOrCreate<EndBorder>();
                return new Border(item);
            }
        }

        public Border InsideHorizontalBorder
        {
            get
            {
                var item = parent.GetOrCreate<InsideHorizontalBorder>();
                return new Border(item);
            }
        }

        public Border InsideVerticalBorder
        {
            get
            {
                var item = parent.GetOrCreate<InsideVerticalBorder>();
                return new Border(item);
            }
        }

        public Border LeftBorder
        {
            get
            {
                var item = parent.GetOrCreate<LeftBorder>();
                return new Border(item);
            }
        }

        public Border RightBorder
        {
            get
            {
                var item = parent.GetOrCreate<RightBorder>();
                return new Border(item);
            }
        }

        public Border StartBorder
        {
            get
            {
                var item = parent.GetOrCreate<StartBorder>();
                return new Border(item);
            }
        }

        public Border TopBorder
        {
            get
            {
                var item = parent.GetOrCreate<TopBorder>();
                return new Border(item);
            }
        }

        public Border TopLeftToBottomRightCellBorder
        {
            get
            {
                var item = parent.GetOrCreate<TopLeftToBottomRightCellBorder>();
                return new Border(item);
            }
        }

        public Border TopRightToBottomLeftCellBorder
        {
            get
            {
                var item = parent.GetOrCreate<TopRightToBottomLeftCellBorder>();
                return new Border(item);
            }
        }

        /// <summary>
        /// Sets the Top, Bottom, Left and Right borders.
        /// </summary>
        /// <param name="size"></param>
        /// <param name="value"></param>
        /// <param name="color"></param>
        public void Set(UInt32Value size, BorderValues value, string color = "auto")
        {
            TopBorder.Set(size, value, color);
            BottomBorder.Set(size, value, color);
            LeftBorder.Set(size, value, color);
            RightBorder.Set(size, value, color);
        }
    }
}
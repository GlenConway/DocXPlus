using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocXPlus
{
    /// <summary>
    ///
    /// </summary>
    public class Borders
    {
        private OpenXmlCompositeElement parent;

        internal Borders(OpenXmlCompositeElement parent) : this(parent, true)
        {
        }

        internal Borders(OpenXmlCompositeElement parent, bool extended)
        {
            this.parent = parent;

            TopBorder.Value = BorderValue.Nil;
            LeftBorder.Value = BorderValue.Nil;
            BottomBorder.Value = BorderValue.Nil;
            RightBorder.Value = BorderValue.Nil;
            InsideHorizontalBorder.Value = BorderValue.Nil;
            InsideVerticalBorder.Value = BorderValue.Nil;

            if (!extended)
            {
                return;
            }

            TopLeftToBottomRightCellBorder.Value = BorderValue.Nil;
            TopRightToBottomLeftCellBorder.Value = BorderValue.Nil;
        }

        /// <summary>
        /// Bottom border
        /// </summary>
        public Border BottomBorder
        {
            get
            {
                var item = parent.GetOrCreate<BottomBorder>();
                return new Border(item);
            }
        }

        /// <summary>
        /// End border
        /// </summary>
        public Border EndBorder
        {
            get
            {
                var item = parent.GetOrCreate<EndBorder>();
                return new Border(item);
            }
        }

        /// <summary>
        /// Inside horizontal border
        /// </summary>
        public Border InsideHorizontalBorder
        {
            get
            {
                var item = parent.GetOrCreate<InsideHorizontalBorder>();
                return new Border(item);
            }
        }

        /// <summary>
        /// Inside vertical border
        /// </summary>
        public Border InsideVerticalBorder
        {
            get
            {
                var item = parent.GetOrCreate<InsideVerticalBorder>();
                return new Border(item);
            }
        }

        /// <summary>
        /// Left border
        /// </summary>
        public Border LeftBorder
        {
            get
            {
                var item = parent.GetOrCreate<LeftBorder>();
                return new Border(item);
            }
        }

        /// <summary>
        /// Right border
        /// </summary>
        public Border RightBorder
        {
            get
            {
                var item = parent.GetOrCreate<RightBorder>();
                return new Border(item);
            }
        }

        /// <summary>
        /// Start border
        /// </summary>
        public Border StartBorder
        {
            get
            {
                var item = parent.GetOrCreate<StartBorder>();
                return new Border(item);
            }
        }

        /// <summary>
        /// Top border
        /// </summary>
        public Border TopBorder
        {
            get
            {
                var item = parent.GetOrCreate<TopBorder>();
                return new Border(item);
            }
        }

        /// <summary>
        /// Top left to bottom right cell border
        /// </summary>
        public Border TopLeftToBottomRightCellBorder
        {
            get
            {
                var item = parent.GetOrCreate<TopLeftToBottomRightCellBorder>();
                return new Border(item);
            }
        }

        /// <summary>
        /// Top right to bottom left cell border
        /// </summary>
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
        /// <param name="size">The size of the border in Twips</param>
        /// <param name="value"></param>
        /// <param name="color"></param>
        public void Set(UInt32Value size, BorderValue value, string color = "auto")
        {
            TopBorder.Set(size, value, color);
            LeftBorder.Set(size, value, color);
            BottomBorder.Set(size, value, color);
            RightBorder.Set(size, value, color);
        }
    }
}
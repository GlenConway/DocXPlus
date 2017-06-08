using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocXPlus
{
    /// <summary>
    /// Represents a border
    /// </summary>
    public class Border
    {
        private BorderType parent;

        internal Border(BorderType parent)
        {
            this.parent = parent;
        }

        /// <summary>
        /// The border color Hex value
        /// </summary>
        public string Color
        {
            get
            {
                return parent.Color;
            }
            set
            {
                parent.Color = value;
            }
        }

        /// <summary>
        /// Specifies if the border has a frame
        /// </summary>
        public bool Frame
        {
            get
            {
                return parent.Frame;
            }
            set
            {
                parent.Frame = value;
            }
        }

        /// <summary>
        /// Specifies if the border has a shadow
        /// </summary>
        public bool Shadow
        {
            get
            {
                return parent.Shadow;
            }
            set
            {
                parent.Shadow = value;
            }
        }

        /// <summary>
        /// The size of the border in Twips
        /// </summary>
        public UInt32Value Size
        {
            get
            {
                return parent.Size;
            }
            set
            {
                parent.Size = value;
            }
        }

        /// <summary>
        /// The spacing of the border in Twips
        /// </summary>
        public UInt32Value Space
        {
            get
            {
                return parent.Space;
            }
            set
            {
                parent.Space = value;
            }
        }

        /// <summary>
        /// The border theme color value
        /// </summary>
        public ThemeColorValues ThemeColor
        {
            get
            {
                return parent.ThemeColor;
            }
            set
            {
                parent.ThemeColor = value;
            }
        }

        /// <summary>
        /// The border theme shade
        /// </summary>
        public string ThemeShade
        {
            get
            {
                return parent.ThemeShade;
            }
            set
            {
                parent.ThemeShade = value;
            }
        }

        /// <summary>
        /// the border theme tint
        /// </summary>
        public string ThemeTint
        {
            get
            {
                return parent.ThemeTint;
            }
            set
            {
                parent.ThemeTint = value;
            }
        }

        /// <summary>
        /// The type of border
        /// </summary>
        public BorderValue Value
        {
            get
            {
                return Convert.ToBorderValue(parent.Val);
            }
            set
            {
                parent.Val = Convert.ToBorderValues(value);
            }
        }

        /// <summary>
        /// Sets the border values
        /// </summary>
        /// <param name="size">The size of the border in Twips</param>
        /// <param name="value"></param>
        /// <param name="color"></param>
        public Border Set(UInt32Value size, BorderValue value, string color = "auto")
        {
            Size = size;
            Value = value;
            Color = color;

            return this;
        }
    }
}
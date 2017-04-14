using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocXPlus
{
    public class Border
    {
        private BorderType parent;

        internal Border(BorderType parent)
        {
            this.parent = parent;
        }

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

        public BorderValues Value
        {
            get
            {
                return parent.Val;
            }
            set
            {
                parent.Val = value;
            }
        }

        /// <summary>
        /// Sets the border values
        /// </summary>
        /// <param name="size">The size of the border in Twips</param>
        /// <param name="value"></param>
        /// <param name="color"></param>
        public void Set(UInt32Value size, BorderValues value, string color = "auto")
        {
            Size = size;
            Value = value;
            Color = color;
        }
    }
}
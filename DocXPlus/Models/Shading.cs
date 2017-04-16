using DocumentFormat.OpenXml.Wordprocessing;

namespace DocXPlus
{
    /// <summary>
    /// Represents shading
    /// </summary>
    public class Shading
    {
        private DocumentFormat.OpenXml.Wordprocessing.Shading shading;

        internal Shading(DocumentFormat.OpenXml.Wordprocessing.Shading shading)
        {
            this.shading = shading;
        }

        /// <summary>
        /// Shading Hex color
        /// </summary>
        public string Color
        {
            get
            {
                return shading.Color;
            }
            set
            {
                shading.Color = value;
            }
        }

        /// <summary>
        /// Shading fill
        /// </summary>
        public string Fill
        {
            get
            {
                return shading.Fill;
            }
            set
            {
                shading.Fill = value;
            }
        }

        /// <summary>
        /// Shading theme color value
        /// </summary>
        public ThemeColorValues ThemeColor
        {
            get
            {
                return shading.ThemeColor;
            }
            set
            {
                shading.ThemeColor = value;
            }
        }

        /// <summary>
        /// Shading theme fill value
        /// </summary>
        public ThemeColorValues ThemeFill
        {
            get
            {
                return shading.ThemeFill;
            }
            set
            {
                shading.ThemeFill = value;
            }
        }

        /// <summary>
        /// Shading theme fill shade
        /// </summary>
        public string ThemeFillShade
        {
            get
            {
                return shading.ThemeFillShade;
            }
            set
            {
                shading.ThemeFillShade = value;
            }
        }

        /// <summary>
        /// Shading theme fill tinit
        /// </summary>
        public string ThemeFillTint
        {
            get
            {
                return shading.ThemeFillTint;
            }
            set
            {
                shading.ThemeFillTint = value;
            }
        }

        /// <summary>
        /// Shading theme shade
        /// </summary>
        public string ThemeShade
        {
            get
            {
                return shading.ThemeShade;
            }
            set
            {
                shading.ThemeShade = value;
            }
        }

        /// <summary>
        /// Shading theme tint
        /// </summary>
        public string ThemeTint
        {
            get
            {
                return shading.ThemeTint;
            }
            set
            {
                shading.ThemeTint = value;
            }
        }

        /// <summary>
        /// Shading pattern
        /// </summary>
        public ShadingPatternValue Val
        {
            get
            {
                return Convert.ToShadingPatternValue(shading.Val);
            }
            set
            {
                shading.Val = Convert.ToShadingPatternValues(value);
            }
        }

        /// <summary>
        /// Sets the pattern, fill and color
        /// </summary>
        /// <param name="value"></param>
        /// <param name="fill">Hex fill color</param>
        /// <param name="color"></param>
        public void Set(ShadingPatternValue value, string fill, string color = "auto")
        {
            Val = value;
            Color = color;
            Fill = fill;
        }
    }
}
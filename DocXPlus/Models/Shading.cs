using DocumentFormat.OpenXml.Wordprocessing;

namespace DocXPlus.Models
{
    public class Shading
    {
        private DocumentFormat.OpenXml.Wordprocessing.Shading shading;

        internal Shading(DocumentFormat.OpenXml.Wordprocessing.Shading shading)
        {
            this.shading = shading;
        }

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

        public ShadingPatternValues Val
        {
            get
            {
                return shading.Val;
            }
            set
            {
                shading.Val = value;
            }
        }

        public void Set(ShadingPatternValues value, string fill, string color = "auto")
        {
            Val = value;
            Color = color;
            Fill = fill;
        }
    }
}
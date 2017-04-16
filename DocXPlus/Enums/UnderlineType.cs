using System;
using System.Collections.Generic;
using System.Text;

namespace DocXPlus
{
    /// <summary>
    /// Types of underline
    /// </summary>
    public enum UnderlineType
    {
        /// <summary>
        /// Single Underline.
        /// </summary>
        Single = 0,
        /// <summary>
        /// Underline Non-Space Characters Only.
        /// </summary>
        Words = 1,
        /// <summary>
        /// Double Underline.
        /// </summary>
        Double = 2,
        /// <summary>
        /// Thick Underline.
        /// </summary>
        Thick = 3,
        /// <summary>
        /// Dotted Underline.
        /// </summary>
        Dotted = 4,
        /// <summary>
        /// Thick Dotted Underline.
        /// </summary>
        DottedHeavy = 5,
        /// <summary>
        /// Dashed Underline.
        /// </summary>
        Dash = 6,
        /// <summary>
        /// Thick Dashed Underline.
        /// </summary>
        DashedHeavy = 7,
        /// <summary>
        /// Long Dashed Underline.
        /// </summary>
        DashLong = 8,
        /// <summary>
        /// Thick Long Dashed Underline.
        /// </summary>
        DashLongHeavy = 9,
        /// <summary>
        /// Dash-Dot Underline.
        /// </summary>
        DotDash = 10,
        /// <summary>
        /// Thick Dash-Dot Underline.
        /// </summary>
        DashDotHeavy = 11,
        /// <summary>
        /// Dash-Dot-Dot Underline.
        /// </summary>
        DotDotDash = 12,
        /// <summary>
        /// Thick Dash-Dot-Dot Underline.
        /// </summary>
        DashDotDotHeavy = 13,
        /// <summary>
        /// Wave Underline.
        /// </summary>
        Wave = 14,
        /// <summary>
        /// Heavy Wave Underline.
        /// </summary>
        WavyHeavy = 15,
        /// <summary>
        /// Double Wave Underline.
        /// </summary>
        WavyDouble = 16,
        /// <summary>
        /// No Underline.
        /// </summary>
        None = 17
    }
}

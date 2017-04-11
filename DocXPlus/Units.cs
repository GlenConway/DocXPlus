using DocumentFormat.OpenXml;

namespace DocXPlus
{
    public static class Units
    {
        internal static Int32Value Inch
        {
            get
            {
                return new Int32Value(1440);
            }
        }

        internal static UInt32Value UHalfInch
        {
            get
            {
                return new UInt32Value((uint)720);
            }
        }

        internal static UInt32Value UInch
        {
            get
            {
                return new UInt32Value((uint)1440);
            }
        }

        internal static UInt32Value UZero
        {
            get
            {
                return new UInt32Value((uint)0);
            }
        }
    }
}
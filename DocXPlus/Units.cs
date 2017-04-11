using DocumentFormat.OpenXml;

namespace DocXPlus
{
    public static class Units
    {
        public static Int32Value Inch
        {
            get
            {
                return new Int32Value(1440);
            }
        }

        public static UInt32Value UHalfInch
        {
            get
            {
                return new UInt32Value((uint)720);
            }
        }

        public static UInt32Value UInch
        {
            get
            {
                return new UInt32Value((uint)1440);
            }
        }

        public static UInt32Value UZero
        {
            get
            {
                return new UInt32Value((uint)0);
            }
        }
    }
}
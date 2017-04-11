using DocumentFormat.OpenXml;

namespace DocXPlus
{
    public static class Units
    {
        public static UInt32Value FourAndHalfPt
        {
            get
            {
                return HalfPt * 9;
            }
        }

        public static UInt32Value HalfPt
        {
            get
            {
                return QtrPt * 2;
            }
        }

        public static Int32Value Inch
        {
            get
            {
                return new Int32Value(1440);
            }
        }

        public static UInt32Value OneAndHalfPt
        {
            get
            {
                return QtrPt * 6;
            }
        }

        public static UInt32Value OnePt
        {
            get
            {
                return QtrPt * 4;
            }
        }

        public static UInt32Value QtrPt
        {
            get
            {
                return new UInt32Value((uint)2);
            }
        }

        public static UInt32Value SixPt
        {
            get
            {
                return ThreePt * 2;
            }
        }

        public static UInt32Value ThreePt
        {
            get
            {
                return OnePt * 3;
            }
        }

        public static UInt32Value ThreeQtrPt
        {
            get
            {
                return QtrPt * 3;
            }
        }

        public static UInt32Value TwoAndQtrPt
        {
            get
            {
                return QtrPt * 9;
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
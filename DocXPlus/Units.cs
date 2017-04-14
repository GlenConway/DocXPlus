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

        /// <summary>
        /// One Inch as Twips
        /// </summary>
        public static Int32Value Inch
        {
            get
            {
                return new Int32Value(InchToTwips(1));
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

        /// <summary>
        /// Half Inch as Twips
        /// </summary>
        public static UInt32Value UHalfInch
        {
            get
            {
                return UInchToTwips(.5);
            }
        }

        /// <summary>
        /// One Inch as Twips
        /// </summary>
        public static UInt32Value UInch
        {
            get
            {
                return UInchToTwips(1);
            }
        }

        /// <summary>
        /// Zero Twips
        /// </summary>
        public static UInt32Value UZero
        {
            get
            {
                return new UInt32Value((uint)0);
            }
        }

        /// <summary>
        /// Converts CM to an English Metric Unit (EMU)
        /// </summary>
        /// <param name="inches"></param>
        /// <returns></returns>
        public static Int64Value CMToEMU(double inches)
        {
            return System.Convert.ToInt64(inches * 360000);
        }

        /// <summary>
        /// Converts an inch to an English Metric Unit (EMU)
        /// </summary>
        /// <param name="inches"></param>
        /// <returns></returns>
        public static Int64Value InchToEMU(double inches)
        {
            return System.Convert.ToInt64(inches * 914400);
        }

        /// <summary>
        /// Converts an inch value such as 1.25 into Twips
        /// </summary>
        /// <param name="inches"></param>
        /// <returns></returns>
        public static Int32Value InchToTwips(double inches)
        {
            return System.Convert.ToInt32(inches * 1440);
        }

        /// <summary>
        /// Converts a points value such as 14 into Twips
        /// </summary>
        /// <param name="points"></param>
        /// <returns></returns>
        public static Int32Value PointsToTwips(int points)
        {
            return points * 20;
        }

        public static UInt32Value UInchToTwips(double inches)
        {
            return System.Convert.ToUInt32(inches * 1440);
        }

        public static UInt32Value UPointsToTwips(int points)
        {
            return System.Convert.ToUInt32(points * 20);
        }
    }
}
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocXPlus
{
    public class TableLook
    {
        private TableProperties tableProperties;

        internal TableLook(TableProperties tableProperties)
        {
            this.tableProperties = tableProperties;
        }

        public bool FirstColumn
        {
            get
            {
                return Look.FirstColumn;
            }
            set
            {
                Look.FirstColumn = value;
            }
        }

        public bool FirstRow
        {
            get
            {
                return Look.FirstRow;
            }
            set
            {
                Look.FirstRow = value;
            }
        }

        public bool LastColumn
        {
            get
            {
                return Look.LastColumn;
            }
            set
            {
                Look.LastColumn = value;
            }
        }

        public bool LastRow
        {
            get
            {
                return Look.LastRow;
            }
            set
            {
                Look.LastRow = value;
            }
        }

        public bool NoHorizontalBand
        {
            get
            {
                return Look.NoHorizontalBand;
            }
            set
            {
                Look.NoHorizontalBand = value;
            }
        }

        public bool NoVerticalBand
        {
            get
            {
                return Look.NoVerticalBand;
            }
            set
            {
                Look.NoVerticalBand = value;
            }
        }

        public string Value
        {
            get
            {
                return Look.Val;
            }
            set
            {
                Look.Val = value;
            }
        }

        private DocumentFormat.OpenXml.Wordprocessing.TableLook Look => tableProperties.GetOrCreate<DocumentFormat.OpenXml.Wordprocessing.TableLook>();
    }
}
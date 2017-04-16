namespace DocXPlus
{
    internal static class Convert
    {
        internal static Align ToAlign(DocumentFormat.OpenXml.Wordprocessing.JustificationValues value)
        {
            switch (value)
            {
                case DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Both:
                    return Align.Both;

                case DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center:
                    return Align.Center;

                case DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Distribute:
                    return Align.Distribute;

                case DocumentFormat.OpenXml.Wordprocessing.JustificationValues.End:
                    return Align.End;

                case DocumentFormat.OpenXml.Wordprocessing.JustificationValues.HighKashida:
                    return Align.HighKashida;

                case DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Left:
                    return Align.Left;

                case DocumentFormat.OpenXml.Wordprocessing.JustificationValues.LowKashida:
                    return Align.LowKashida;

                case DocumentFormat.OpenXml.Wordprocessing.JustificationValues.MediumKashida:
                    return Align.MediumKashida;

                case DocumentFormat.OpenXml.Wordprocessing.JustificationValues.NumTab:
                    return Align.NumTab;

                case DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Right:
                    return Align.Right;

                case DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Start:
                    return Align.Start;

                case DocumentFormat.OpenXml.Wordprocessing.JustificationValues.ThaiDistribute:
                    return Align.ThaiDistribute;
            }

            return Align.Left;
        }

        internal static DocumentType ToDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType value)
        {
            return (DocumentType)((int)value);
        }

        internal static DocumentFormat.OpenXml.Wordprocessing.JustificationValues ToJustificationValues(Align value)
        {
            switch (value)
            {
                case Align.Both:
                    return DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Both;

                case Align.Center:
                    return DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center;

                case Align.Distribute:
                    return DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Distribute;

                case Align.End:
                    return DocumentFormat.OpenXml.Wordprocessing.JustificationValues.End;

                case Align.HighKashida:
                    return DocumentFormat.OpenXml.Wordprocessing.JustificationValues.HighKashida;

                case Align.Left:
                    return DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Left;

                case Align.LowKashida:
                    return DocumentFormat.OpenXml.Wordprocessing.JustificationValues.LowKashida;

                case Align.MediumKashida:
                    return DocumentFormat.OpenXml.Wordprocessing.JustificationValues.MediumKashida;

                case Align.NumTab:
                    return DocumentFormat.OpenXml.Wordprocessing.JustificationValues.NumTab;

                case Align.Right:
                    return DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Right;

                case Align.Start:
                    return DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Start;

                case Align.ThaiDistribute:
                    return DocumentFormat.OpenXml.Wordprocessing.JustificationValues.ThaiDistribute;
            }

            return DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Left;
        }

        internal static PageOrientation ToPageOrientation(DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues value)
        {
            return (PageOrientation)((int)value);
        }

        internal static DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues ToPageOrientationValues(PageOrientation value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues)((int)value);
        }

        internal static UnderlineType ToUnderlineType(DocumentFormat.OpenXml.Wordprocessing.UnderlineValues value)
        {
            return (UnderlineType)((int)value);
        }

        internal static DocumentFormat.OpenXml.Wordprocessing.UnderlineValues ToUnderlineValues(UnderlineType value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.UnderlineValues)((int)value);
        }

        internal static DocumentFormat.OpenXml.WordprocessingDocumentType ToWordprocessingDocumentType(DocumentType value)
        {
            return (DocumentFormat.OpenXml.WordprocessingDocumentType)((int)value);
        }
    }
}
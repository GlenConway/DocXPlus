namespace DocXPlus
{
    internal static class Utils
    {
        internal static Align ConvertToAlign(DocumentFormat.OpenXml.Wordprocessing.JustificationValues value)
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

        internal static DocumentFormat.OpenXml.Wordprocessing.JustificationValues ConvertToJustificationValues(Align value)
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
    }
}
namespace DocXPlus
{
    internal static class Convert
    {
        internal static Align ToAlign(DocumentFormat.OpenXml.Wordprocessing.JustificationValues value)
        {
            return (Align)((int)value);
        }

        internal static BorderValue ToBorderValue(DocumentFormat.OpenXml.Wordprocessing.BorderValues value)
        {
            return (BorderValue)((int)value);
        }

        internal static DocumentFormat.OpenXml.Wordprocessing.BorderValues ToBorderValues(BorderValue value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.BorderValues)((int)value);
        }

        internal static DocumentType ToDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType value)
        {
            return (DocumentType)((int)value);
        }

        internal static DocumentFormat.OpenXml.Wordprocessing.JustificationValues ToJustificationValues(Align value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.JustificationValues)((int)value);
        }

        internal static PageOrientation ToPageOrientation(DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues value)
        {
            return (PageOrientation)((int)value);
        }

        internal static DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues ToPageOrientationValues(PageOrientation value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues)((int)value);
        }

        internal static ShadingPattern ToShadingPattern(DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues value)
        {
            return (ShadingPattern)((int)value);
        }

        internal static DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues ToShadingPatternValues(ShadingPattern value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues)((int)value);
        }

        internal static StyleValue ToStyleValue(DocumentFormat.OpenXml.Wordprocessing.StyleValues value)
        {
            return (StyleValue)((int)value);
        }

        internal static DocumentFormat.OpenXml.Wordprocessing.StyleValues ToStyleValues(StyleValue value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.StyleValues)((int)value);
        }

        internal static TableVerticalAlignment ToTableVerticalAlignment(DocumentFormat.OpenXml.Wordprocessing.TableVerticalAlignmentValues value)
        {
            return (TableVerticalAlignment)((int)value);
        }

        internal static DocumentFormat.OpenXml.Wordprocessing.TableVerticalAlignmentValues ToTableVerticalAlignmentValues(TableVerticalAlignment value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.TableVerticalAlignmentValues)((int)value);
        }

        internal static TableWidthUnitValue ToTableWidthUnitValue(DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues value)
        {
            return (TableWidthUnitValue)((int)value);
        }

        internal static DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues ToTableWidthUnitValues(TableWidthUnitValue value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues)((int)value);
        }

        internal static TableWidthValue ToTableWidthValue(DocumentFormat.OpenXml.Wordprocessing.TableWidthValues value)
        {
            return (TableWidthValue)((int)value);
        }

        internal static DocumentFormat.OpenXml.Wordprocessing.TableWidthValues ToTableWidthValues(TableWidthValue value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.TableWidthValues)((int)value);
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
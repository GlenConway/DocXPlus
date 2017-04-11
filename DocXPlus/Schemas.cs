using DocumentFormat.OpenXml;

namespace DocXPlus
{
    public static class Schemas
    {
        public const string m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        public const string mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        public const string o = "urn:schemas-microsoft-com:office:office";
        public const string r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public const string v = "urn:schemas-microsoft-com:vml";
        public const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public const string w10 = "urn:schemas-microsoft-com:office:word";
        public const string w14 = "http://schemas.microsoft.com/office/word/2010/wordml";
        public const string wne = "http://schemas.microsoft.com/office/word/2006/wordml";
        public const string wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        public const string wp14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
        public const string wpc = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
        public const string wpg = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
        public const string wpi = "http://schemas.microsoft.com/office/word/2010/wordprocessingInk";
        public const string wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

        public static void AddNamespaceDeclarations(OpenXmlPartRootElement element)
        {
            element.AddNamespaceDeclaration("wpc", wpc);
            element.AddNamespaceDeclaration("mc", mc);
            element.AddNamespaceDeclaration("o", o);
            element.AddNamespaceDeclaration("r", r);
            element.AddNamespaceDeclaration("m", m);
            element.AddNamespaceDeclaration("v", v);
            element.AddNamespaceDeclaration("wp14", wp14);
            element.AddNamespaceDeclaration("wp", wp);
            element.AddNamespaceDeclaration("w10", w10);
            element.AddNamespaceDeclaration("w", w);
            element.AddNamespaceDeclaration("w14", w14);
            element.AddNamespaceDeclaration("wpg", wpg);
            element.AddNamespaceDeclaration("wpi", wpi);
            element.AddNamespaceDeclaration("wne", wne);
            element.AddNamespaceDeclaration("wps", wps);
        }
    }
}
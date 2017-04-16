using DocXPlus;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace DocXPlusTests
{
    [TestClass]
    public class ImageTests : TestBase
    {
        [TestMethod]
        public void AddSingleImage()
        {
            var filename = Path.Combine(TempDirectory, "AddSingleImage.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            var width = 0d;
            var height = 0d;

            using (System.Drawing.Image img = System.Drawing.Image.FromStream(GetEmbeddedResourceStream("Peggys_Cove_Nova_Scotia_01.jpg"), useEmbeddedColorManagement: false, validateImageData: false))
            {
                width = img.Width;
                height = img.Height;
            }

            double ratio = height / width;

            var inchWidth = 2;
            var inchHeight = inchWidth * ratio;

            var drawing = doc.AddImage(GetEmbeddedResourceStream("Peggys_Cove_Nova_Scotia_01.jpg"), "image/jpg", Units.InchToEMU(inchWidth), Units.InchToEMU(inchHeight));
            doc.AddParagraph().Append(drawing);

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }

        [TestMethod]
        public void AddTwoImages()
        {
            var filename = Path.Combine(TempDirectory, "AddTwoImages.docx");

            var doc = DocX.Create(filename, DocumentType.Document);

            var width = 0d;
            var height = 0d;

            using (System.Drawing.Image img = System.Drawing.Image.FromStream(GetEmbeddedResourceStream("Peggys_Cove_Nova_Scotia_01.jpg"), useEmbeddedColorManagement: false, validateImageData: false))
            {
                width = img.Width;
                height = img.Height;
            }

            double ratio = height / width;

            var inchWidth = 2;
            var inchHeight = inchWidth * ratio;

            var drawing = doc.AddImage(GetEmbeddedResourceStream("Peggys_Cove_Nova_Scotia_01.jpg"), "image/jpg", Units.InchToEMU(inchWidth), Units.InchToEMU(inchHeight));
            doc.AddParagraph().Append(drawing);

            drawing = doc.AddImage(GetEmbeddedResourceStream("Peggys_Cove_Nova_Scotia_01.jpg"), "image/jpg", Units.InchToEMU(inchWidth * 2), Units.InchToEMU(inchHeight * 2));
            doc.AddParagraph().Append(drawing);

            doc.Close();

            ValidateWordDocument(filename);

            Launch(filename);
        }
    }
}
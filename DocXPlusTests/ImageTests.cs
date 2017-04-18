using DocXPlus;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace DocXPlusTests
{
    [TestClass]
    public class ImageTests : TestBase
    {
        [TestMethod]
        public void AddBmp()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                var width = 0d;
                var height = 0d;

                using (System.Drawing.Image img = System.Drawing.Image.FromStream(GetEmbeddedResourceStream("Peggys_Cove_Nova_Scotia_01.bmp"), useEmbeddedColorManagement: false, validateImageData: false))
                {
                    width = img.Width;
                    height = img.Height;
                }

                double ratio = height / width;

                var inchWidth = 2;
                var inchHeight = inchWidth * ratio;

                var drawing = doc.AddImage(GetEmbeddedResourceStream("Peggys_Cove_Nova_Scotia_01.bmp"), "image/bmp", Units.InchToEMU(inchWidth), Units.InchToEMU(inchHeight));
                doc.AddParagraph().Append(drawing);

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void AddByteArray()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                var width = 0d;
                var height = 0d;

                byte[] bytes = null;

                using (var stream = new MemoryStream())
                {
                    GetEmbeddedResourceStream("Peggys_Cove_Nova_Scotia_01.jpg").CopyTo(stream);

                    bytes = stream.ToArray();
                }

                using (System.Drawing.Image img = System.Drawing.Image.FromStream(GetEmbeddedResourceStream("Peggys_Cove_Nova_Scotia_01.jpg"), useEmbeddedColorManagement: false, validateImageData: false))
                {
                    width = img.Width;
                    height = img.Height;
                }

                double ratio = height / width;

                var inchWidth = 2;
                var inchHeight = inchWidth * ratio;

                var drawing = doc.AddImage(bytes, "image/jpg", Units.InchToEMU(inchWidth), Units.InchToEMU(inchHeight));
                doc.AddParagraph().Append(drawing);

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void AddPng()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                var width = 0d;
                var height = 0d;

                using (System.Drawing.Image img = System.Drawing.Image.FromStream(GetEmbeddedResourceStream("Peggys_Cove_Nova_Scotia_01.png"), useEmbeddedColorManagement: false, validateImageData: false))
                {
                    width = img.Width;
                    height = img.Height;
                }

                double ratio = height / width;

                var inchWidth = 2;
                var inchHeight = inchWidth * ratio;

                var drawing = doc.AddImage(GetEmbeddedResourceStream("Peggys_Cove_Nova_Scotia_01.png"), "image/png", Units.InchToEMU(inchWidth), Units.InchToEMU(inchHeight));
                doc.AddParagraph().Append(drawing);

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void AddSingleImage()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void AddTwoImages()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void InFooter()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddFooters();

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

                var drawing = doc.DefaultFooter.AddImage(GetEmbeddedResourceStream("Peggys_Cove_Nova_Scotia_01.jpg"), "image/jpg", Units.InchToEMU(inchWidth), Units.InchToEMU(inchHeight));
                doc.DefaultFooter.AddParagraph().Append(drawing);

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void InHeader()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.AddHeaders();

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

                var drawing = doc.DefaultHeader.AddImage(GetEmbeddedResourceStream("Peggys_Cove_Nova_Scotia_01.jpg"), "image/jpg", Units.InchToEMU(inchWidth), Units.InchToEMU(inchHeight));
                doc.DefaultHeader.AddParagraph().Append(drawing);

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void WrongContentType()
        {
            using (var doc = new DocX())
            {
                doc.Create();

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

                var drawing = doc.AddImage(GetEmbeddedResourceStream("Peggys_Cove_Nova_Scotia_01.jpg"), "image/bmp", Units.InchToEMU(inchWidth), Units.InchToEMU(inchHeight));
                doc.AddParagraph().Append(drawing);

                Validate(doc);

                doc.Close();
            }
        }
    }
}
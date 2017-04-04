using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;

namespace DocXPlusTests
{
    [TestClass]
    public class DocumentTests
    {
        protected readonly string TempDirectory;

        public DocumentTests()
        {
            TempDirectory = Path.Combine(Path.GetTempPath(), "DocXPlusTests", "documents");
            Setup(TempDirectory);
        }

        public static void ValidateWordDocument(string filepath)
        {
            using (WordprocessingDocument wordprocessingDocument =
            WordprocessingDocument.Open(filepath, true))
            {
                try
                {
                    OpenXmlValidator validator = new OpenXmlValidator();
                    var validation = validator.Validate(wordprocessingDocument);

                    Assert.AreEqual(0, validation.Count());

                //    int count = 0;
                //    foreach (ValidationErrorInfo error in
                //        )
                //    {
                //        count++;
                //        Console.WriteLine("Error " + count);
                //        Console.WriteLine("Description: " + error.Description);
                //        Console.WriteLine("ErrorType: " + error.ErrorType);
                //        Console.WriteLine("Node: " + error.Node);
                //        Console.WriteLine("Path: " + error.Path.XPath);
                //        Console.WriteLine("Part: " + error.Part.Uri);
                //        Console.WriteLine("-------------------------------------------");
                //    }

                //    Console.WriteLine("count={0}", count);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                finally {
                    wordprocessingDocument.Close();
                }
            }
        }

        [TestMethod]
        public void CreateUsingFile()
        {
            var filename = Path.Combine(TempDirectory, "CreateUsingFile.docx");

            var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            ValidateWordDocument(filename);
        }

        private static void Setup(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
    }
}
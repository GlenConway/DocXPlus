using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;

namespace DocXPlusTests
{
    [TestClass]
    public class TestBase
    {
        protected readonly string TempDirectory;

        public TestBase()
        {
            TempDirectory = Path.Combine(Path.GetTempPath(), "DocXPlusTests", "documents");
            Setup(TempDirectory);
        }

        protected void ValidateWordDocument(string filepath)
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
                finally
                {
                    wordprocessingDocument.Close();
                }
            }
        }

        private static void Setup(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

        protected void Launch(string filename)
        {
            System.Diagnostics.Process.Start(filename);
        }
    }
}
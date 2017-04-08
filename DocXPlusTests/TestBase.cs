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

        protected string LoremIpsum
        {
            get
            {
                return "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nulla et scelerisque eros, at posuere nisl. Maecenas pretium porta tellus sit amet pulvinar. Quisque arcu elit, consequat sit amet finibus non, sodales rutrum risus. Nulla sagittis nunc nec auctor rhoncus. Maecenas dictum nunc vel lobortis auctor. Quisque mattis imperdiet mattis. Vivamus lacinia maximus diam sed posuere. Quisque ullamcorper mi quis ipsum condimentum finibus. Nullam mollis sit amet ex ullamcorper venenatis. Pellentesque lacinia porta leo, eget imperdiet mauris.";
            }
        }

        protected void Launch(string filename)
        {
            System.Diagnostics.Process.Start(filename);
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
    }
}
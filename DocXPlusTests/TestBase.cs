using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

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

        protected Stream GetEmbeddedResourceStream(string name)
        {
            var assembly = Assembly.GetExecutingAssembly();

            var resourceName = string.Format("DocXPlusTests.Resources.{0}", name);

            var resourceNames = assembly.GetManifestResourceNames();

            if (!resourceNames.Contains(resourceName))
                throw new InvalidOperationException("Assembly does not contain a resource named '" + resourceName + "'.");

            return assembly.GetManifestResourceStream(resourceName);
        }

        protected void Launch(string filename)
        {
            System.Diagnostics.Process.Start(filename);
        }

        protected void Validate(DocXPlus.DocX document)
        {
            using (var stream = new MemoryStream())
            {
                document.SaveAs(stream);

                using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(stream, true))
                {
                    try
                    {
                        OpenXmlValidator validator = new OpenXmlValidator();
                        var validation = validator.Validate(wordprocessingDocument);

                        var sb = new StringBuilder();

                        foreach (ValidationErrorInfo error in validation)
                        {
                            sb.AppendLine("Description: " + error.Description);
                            sb.AppendLine("ErrorType: " + error.ErrorType);
                            sb.AppendLine("Node: " + error.Node);
                            sb.AppendLine("Path: " + error.Path.XPath);
                            sb.AppendLine("Part: " + error.Part.Uri);

                            sb.AppendLine(string.Empty);
                        }

                        if (validation.Count() > 0)
                        {
                            Assert.Fail(sb.ToString());
                        }
                    }
                    finally
                    {
                        wordprocessingDocument.Close();
                    }
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
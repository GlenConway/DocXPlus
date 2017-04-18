using DocXPlus;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DocXPlusTests
{
    [TestClass]
    public class DocumentTests : TestBase
    {
        [TestMethod]
        public void Portrait()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                Validate(doc);

                doc.Close();
            }
        }

        [TestMethod]
        public void Landscape()
        {
            using (var doc = new DocX())
            {
                doc.Create();

                doc.Orientation = PageOrientation.Landscape;

                Validate(doc);

                doc.Close();
            }
        }
    }
}
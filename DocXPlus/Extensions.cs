using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

namespace DocXPlus
{
    internal static class Extensions
    {
        internal static void Bold(this Run run)
        {
            RunProperties runProperties = run.GetOrCreate<RunProperties>(true);
            Bold prop = runProperties.GetOrCreate<Bold>();
            prop.Val = OnOffValue.FromBoolean(true);
        }

        internal static T GetOrCreate<T>(this OpenXmlCompositeElement element) where T : OpenXmlElement, new()
        {
            return element.GetOrCreate<T>(false);
        }

        internal static T GetOrCreate<T>(this OpenXmlCompositeElement element, bool prepend) where T : OpenXmlElement, new()
        {
            if (!element.Has<T>())
            {
                if (prepend)
                {
                    element.PrependChild(new T());
                }
                else
                {
                    element.AppendChild(new T());
                }
            }

            return element.Elements<T>().First();
        }

        internal static bool Has<T>(this OpenXmlCompositeElement element) where T : OpenXmlElement
        {
            return element.Elements<T>().Count() != 0;
        }

        internal static void Italic(this Run run)
        {
            RunProperties runProperties = run.GetOrCreate<RunProperties>(true);
            Italic prop = runProperties.GetOrCreate<Italic>();
            prop.Val = OnOffValue.FromBoolean(true);
        }

        internal static void Underline(this Run run, UnderlineValues value)
        {
            RunProperties runProperties = run.GetOrCreate<RunProperties>(true);
            Underline prop = runProperties.GetOrCreate<Underline>();
            prop.Val = value;
        }
    }
}
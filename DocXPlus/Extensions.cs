using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;

namespace DocXPlus
{
    internal static class Extensions
    {
        internal static void Bold(this Run run)
        {
            RunProperties runProperties = run.GetOrCreate<RunProperties>(true);
            Bold prop = runProperties.GetOrCreate<Bold>();
            prop.Val = true;
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

        internal static void RemoveAllChildren<T>(this OpenXmlElement value, string localName, string namespaceUri, string match) where T : OpenXmlElement
        {
            if (!value.HasChildren)
                return;

            OpenXmlElement element = value.FirstChild;
            OpenXmlElement next;

            while (element != null)
            {
                next = element.NextSibling();

                if (element is T)
                {
                    if (!element.HasAttributes)
                        continue;

                    var attribute = element.GetAttribute(localName, namespaceUri);

                    if (attribute.Value.Equals(match, StringComparison.OrdinalIgnoreCase))
                        value.RemoveChild(element);
                }

                element = next;
            }
        }

        internal static void Underline(this Run run, UnderlineValues value)
        {
            RunProperties runProperties = run.GetOrCreate<RunProperties>(true);
            Underline prop = runProperties.GetOrCreate<Underline>();
            prop.Val = value;
        }
    }
}
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
            Bold prop = runProperties.GetOrCreate<Bold>(true);
            prop.Val = true;
        }

        internal static OpenXmlElement Find<T>(this OpenXmlCompositeElement element) where T : OpenXmlElement, new()
        {
            if (element.Elements<T>().Count() > 0)
            {
                return element.Elements<T>().First();
            }

            return null;
        }

        internal static void FontFamily(this Run run, string name)
        {
            RunProperties runProperties = run.GetOrCreate<RunProperties>(true);
            RunFonts prop = runProperties.GetOrCreate<RunFonts>();
            prop.Ascii = name;
            prop.HighAnsi = name;
            prop.ComplexScript = name;
            prop.EastAsia = name;
        }

        internal static void FontSize(this Run run, double size)
        {
            RunProperties runProperties = run.GetOrCreate<RunProperties>(true);

            FontSize fontSize = runProperties.GetOrCreate<FontSize>();
            fontSize.Val = size.ToString();

            FontSizeComplexScript fontSizeComplexScript = runProperties.GetOrCreate<FontSizeComplexScript>();
            fontSizeComplexScript.Val = size.ToString();
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

        internal static T GetOrCreateAfter<T>(this OpenXmlCompositeElement element, OpenXmlElement after) where T : OpenXmlElement, new()
        {
            if (!element.Has<T>())
            {
                element.InsertAfter(new T(), after);
            }

            return element.Elements<T>().First();
        }

        internal static T GetOrCreateBefore<T>(this OpenXmlCompositeElement element, OpenXmlElement before) where T : OpenXmlElement, new()
        {
            if (!element.Has<T>())
            {
                element.InsertBefore(new T(), before);
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

        internal static void SetStyle(this Run run, string styleId)
        {
            var runProperties = run.GetOrCreate<RunProperties>(true);
            var style = runProperties.GetOrCreate<RunStyle>();
            style.Val = styleId;
        }

        internal static OnOffOnlyValues ToOnOffOnlyValues(this bool value)
        {
            return value ? OnOffOnlyValues.On : OnOffOnlyValues.Off;
        }

        internal static void Underline(this Run run, UnderlineType value)
        {
            var runProperties = run.GetOrCreate<RunProperties>(true);
            var prop = runProperties.GetOrCreate<Underline>();
            prop.Val = Convert.ToUnderlineValues(value);
        }
    }
}
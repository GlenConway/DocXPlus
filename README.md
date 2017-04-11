# DocXPlus
Similar to [DocX](https://github.com/WordDocX/DocX) but based on the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) instead of directly
working with the underlying XML. The [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) is very powerful but requires a significant amount of coding. This library provides an abstraction layer 
greatly reducing the amount of code you have to write and maintain.

This initial phase is very much a minimum viable product, only implementing features as they are required. Currently there is no support for reading
existing Word documents.

Stay tuned, under development

## Some Code
**Different Page Orientation After Section Break**
``` c#
var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

doc.Orientation = DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues.Landscape;

doc.AddParagraph().Append("Landscape");

doc.InsertSectionPageBreak();

doc.Orientation = DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues.Portrait;

doc.AddParagraph().Append("Portrait");

doc.InsertSectionPageBreak();

doc.Orientation = DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues.Landscape;

doc.AddParagraph().Append("Landscape");

doc.Close();
```
**Different Headers and Footers After Section Break**
```c#
var doc = DocXPlus.DocX.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

doc.AddHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Header 1");
doc.AddFooter(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Footer 1");

doc.InsertSectionPageBreak();

doc.AddHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Header 2");
doc.AddFooter(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Footer 2");

doc.InsertSectionPageBreak();

doc.AddHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Header 3");
doc.AddFooter(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default).AddParagraph().Append("Footer 3");

doc.Close();
```
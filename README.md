# DocXPlus
Similar to [DocX](https://github.com/WordDocX/DocX) but based on the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) instead of directly
working with the underlying XML. The [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) is very powerful but requires a significant amount of coding. This library provides an abstraction layer 
greatly reducing the amount of code you have to write and maintain.

This initial phase is very much a minimum viable product, only implementing features as they are required. Currently there is no support for reading
existing Word documents.

Version 2.7 of the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) is required so follow the [instructions](https://github.com/OfficeDev/Open-XML-SDK#where-to-get-the-nuget-package) to add the MyGet feed.

Stay tuned, under development

## Some Code
**Different Page Orientation After Section Break**
``` c#
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

var doc = DocXPlus.DocX.Create(filename, WordprocessingDocumentType.Document);
doc.Orientation = PageOrientationValues.Landscape;

doc.AddParagraph().Append("Landscape");

doc.InsertSectionPageBreak();
doc.Orientation = PageOrientationValues.Portrait;

doc.AddParagraph().Append("Portrait");

doc.InsertSectionPageBreak();
doc.Orientation = PageOrientationValues.Landscape;

doc.AddParagraph().Append("Landscape");

doc.Close();
```
**Different Headers and Footers After Section Break**
```c#
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

var doc = DocXPlus.DocX.Create(filename, WordprocessingDocumentType.Document);

doc.AddHeaders();
doc.AddFooters();

doc.DefaultHeader.AddParagraph().Append("Header 1");
doc.DefaultFooter.AddParagraph().Append("Footer 1");

doc.InsertSectionPageBreak();

doc.AddHeaders();
doc.AddFooters();

doc.DefaultHeader.AddParagraph().Append("Header 2");
doc.DefaultFooter.AddParagraph().Append("Footer 2");

doc.InsertSectionPageBreak();

doc.AddHeaders();
doc.AddFooters();

doc.DefaultHeader.AddParagraph().Append("Header 3");
doc.DefaultFooter.AddParagraph().Append("Footer 3");

doc.Close();
```
**Tables**
```c#
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocXPlus;

var doc = DocX.Create(filename, WordprocessingDocumentType.Document);

// 5 columns
var table = doc.AddTable(5);

for (int i = 0; i < 5; i++)
{
    var row = table.AddRow();

    row.SetBorders(Units.HalfPt, BorderValues.Single);

    if (i == 0)
    {
        // shade the first row and set as a header
        row.SetShading(ShadingPatternValues.Clear, "E7E6E6");

        row.HeaderRow = true;
    }

    for (int j = 0; j < 5; j++)
    {
        row.Cells[j].Paragraphs[0].Append($"Cell {(j + 1)}");
    }

    if (i > 0)
    {
        row.Cells[0].MergeRight = i;
    }
}

doc.Close();
```
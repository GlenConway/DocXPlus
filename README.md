# DocXPlus
Similar to [DocX](https://github.com/WordDocX/DocX) but based on the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) instead of directly
working with the underlying XML. The [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) is very powerful but requires a significant amount of coding. This library provides an abstraction layer 
greatly reducing the amount of code you have to write and maintain.

This initial phase is very much a minimum viable product, only implementing features as they are required. Currently there is no support for reading
existing Word documents.

DocXPlus targets .NET Standard 1.3 so version 2.7 of the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) with .NET Standard support is required> Currently this is only available through MyGet so follow these [instructions](https://github.com/OfficeDev/Open-XML-SDK#where-to-get-the-nuget-package) to add the Open XML SDK MyGet feed.

Stay tuned, under development

## Some Code
**Create Using Default Stream**
``` c#
using (var doc = new DocX())
{
    doc.Create();

    doc.AddParagraph().Append(LoremIpsum);

    using (var stream = new FileStream(filename, FileMode.Create))
    {
        doc.SaveAs(stream);
    }
}
```

**Different Page Orientation After Section Break**
``` c#
var doc = DocX.Create(filename, DocumentType.Document);
doc.Orientation = PageOrientation.Landscape;

doc.AddParagraph().Append("Landscape");

doc.InsertSectionPageBreak();
doc.Orientation = PageOrientation.Portrait;

doc.AddParagraph().Append("Portrait");

doc.InsertSectionPageBreak();
doc.Orientation = PageOrientation.Landscape;

doc.AddParagraph().Append("Landscape");

doc.Close();
```
**Different Headers and Footers After Section Break**
```c#
var doc = DocX.Create(filename, DocumentType.Document);

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
You can also add page numbers to the footer
```c#
doc.AddFooters();

doc.DefaultFooter
    .AddParagraph()
    .Append("Page ")
    .AppendPageNumber(PageNumberFormat.Normal)
    .Append(" of ")
    .AppendPageCount(PageNumberFormat.Normal)
    .Bold()
    .Alignment = Align.Center;
```
**Tables**
```c#
 var doc = DocX.Create(filename, DocumentType.Document);

var table = doc.AddTable(5);

for (int i = 0; i < 50; i++)
{
    var row = table.AddRow();
    row.SetBorders(Units.HalfPt, BorderValue.Single);

    if (i == 0)
    {
        // shade the first row and set as a header
        row.SetShading("E7E6E6");

        row.HeaderRow = true;
    }

    for (int j = 0; j < 5; j++)
    {
        row.Cells[j].Paragraphs[0].Append($"Cell {(j + 1)}");
    }
}

doc.Close();
```
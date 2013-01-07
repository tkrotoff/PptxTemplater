# C# library to generate PowerPoint files from templates

This library uses the Office Open XML format (pptx) through the [Open XML SDK 2.0 for Microsoft Office](http://www.microsoft.com/en-us/download/details.aspx?id=5124).
Generated files should be opened using Microsoft PowerPoint >= 2010.

PptxTemplater handles:
- Text tags
- Slides (add/remove)
- Slide notes
- Tables (add/remove columns)
- Pictures

## Example

Create a PowerPoint template with two slides and inserts tags (`{{hello}}`, `{{bonjour}}`, `{{hola}}`) in it,
then generate the final PowerPoint file using the following code:

```C#
const string srcFileName = "template.pptx";
const string dstFileName = "final.pptx";
File.Delete(dstFileName);
File.Copy(srcFileName, dstFileName);

Pptx pptx = new Pptx(dstFileName, FileAccess.ReadWrite);
int nbSlides = pptx.SlidesCount();
Assert.AreEqual(2, nbSlides);

// First slide
{
    PptxSlide slide = pptx.GetSlide(0);
    slide.ReplaceTag("{{hello}}", "HELLO HOW ARE YOU?", PptxSlide.ReplacementType.Global);
    slide.ReplaceTag("{{bonjour}}", "BONJOUR TOUT LE MONDE", PptxSlide.ReplacementType.Global);
    slide.ReplaceTag("{{hola}}", "HOLA MAMA QUE TAL?", PptxSlide.ReplacementType.Global);
}

// Second slide
{
    PptxSlide slide = pptx.GetSlide(1);
    slide.ReplaceTag("{{hello}}", "H", PptxSlide.ReplacementType.Global);
    slide.ReplaceTag("{{bonjour}}", "B", PptxSlide.ReplacementType.Global);
    slide.ReplaceTag("{{hola}}", "H", PptxSlide.ReplacementType.Global);
}
```

## Implementation

The source code is clean, documented, tested and should be stable.
A good amount of unit tests come with the source code.

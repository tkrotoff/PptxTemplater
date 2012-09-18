namespace PptxTemplater
{
    using System.Linq;

    using A = DocumentFormat.OpenXml.Drawing;

    /// <summary>
    /// Represents a table inside a PowerPoint file.
    /// </summary>
    /// <remarks>Could not simply be named Table, conflicts with DocumentFormat.OpenXml.Drawing.Table.</remarks>
    public class PptxTable
    {
        private PptxSlide slideTemplate;
        private readonly int tblId;

        internal PptxTable(PptxSlide slideTemplate, int tblId)
        {
            this.slideTemplate = slideTemplate;
            this.tblId = tblId;
        }

        public class Cell
        {
            internal string Tag { get; set; }

            internal string NewText { get; set; }

            public Cell(string tag, string newText)
            {
                this.Tag = tag;
                this.NewText = newText;
            }
        }

        public void SetRows(params Cell[][] rows)
        {
            // TODO throw an exception if this method is being called several times for the same table

            // Create a new slide from the template slide
            PptxSlide slide = this.slideTemplate.Clone();
            this.slideTemplate.InsertAfter(slide);
            A.Table tbl = PptxSlide.FindTable(slide, this.tblId);

            // donePerSlide starts at 1 instead of 0 because we don't care about the first row
            // The first row contains the titles for the columns
            for (int i = 0, donePerSlide = 1; i < rows.Count();)
            {
                Cell[] row = rows[i];

                if (donePerSlide < RowsCount(tbl))
                {
                    A.TableRow tr = GetRow(tbl, donePerSlide);

                    foreach (A.Paragraph p in tr.Descendants<A.Paragraph>())
                    {
                        foreach (Cell cell in row)
                        {
                            PptxParagraph.ReplaceTag(p, cell.Tag, cell.NewText);
                        }
                    }

                    i++;
                    donePerSlide++;
                }
                else
                {
                    // Create a new slide since the current one is "full"
                    PptxSlide newSlide = this.slideTemplate.Clone();
                    this.slideTemplate.InsertAfter(newSlide);
                    tbl = PptxSlide.FindTable(newSlide, this.tblId);

                    // Not modifying i
                    donePerSlide = 1;
                }
            }

            // Delete the template slide
            this.slideTemplate.Delete();
        }

        private static int RowsCount(A.Table tbl)
        {
            return tbl.Descendants<A.TableRow>().Count();
        }

        private static A.TableRow GetRow(A.Table tbl, int row)
        {
            A.TableRow tr = tbl.Descendants<A.TableRow>().ElementAt(row);
            return tr;
        }
    }
}

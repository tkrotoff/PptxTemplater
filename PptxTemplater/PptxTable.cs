namespace PptxTemplater
{
    using System.Collections.Generic;
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

        /// <summary>
        /// Represents a cell inside a table.
        /// </summary>
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

        /// <summary>
        /// Changes the cells from the table.
        /// </summary>
        /// <remarks>
        /// This method should be called only once.
        /// This method can potentially change the number of slides (by inserting new slides) so you are better off
        /// calling it last.
        /// </remarks>
        public void SetRows(IList<Cell[]> rows)
        {
            // TODO throw an exception if this method is being called several times for the same table

            // Create a new slide from the template slide
            PptxSlide slide = this.slideTemplate.Clone();
            this.slideTemplate.InsertAfter(slide);
            A.Table tbl = slide.FindTable(this.tblId);

            // donePerSlide starts at 1 instead of 0 because we don't care about the first row
            // The first row contains the titles for the columns
            int donePerSlide = 1;
            for (int i = 0; i < rows.Count(); )
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

                    donePerSlide++;
                    i++;
                }
                else
                {
                    // Save the previous slide
                    slide.Save();

                    // Create a new slide since the current one is "full"
                    PptxSlide newSlide = this.slideTemplate.Clone();
                    slide.InsertAfter(newSlide);
                    tbl = newSlide.FindTable(this.tblId);
                    slide = newSlide;

                    donePerSlide = 1;
                    // Not modifying i => do the replacement with the new slide
                }
            }

            // Remove the last remaining rows if any
            for (int row = RowsCount(tbl) - 1; row >= donePerSlide; row--)
            {
                A.TableRow tr = GetRow(tbl, row);
                tr.Remove();
            }

            // Save the latest slide
            // Mandatory otherwise the next time SetRows() is run (on a different table)
            // the rows from the previous tables will not contained the right data (from PptxParagraph.ReplaceTag())
            slide.Save();

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

namespace PptxTemplater
{
    using System.Linq;

    using DocumentFormat.OpenXml.Packaging;
    using A = DocumentFormat.OpenXml.Drawing;

    /// <summary>
    /// Represents a table inside a PowerPoint file.
    /// </summary>
    /// <remarks>Could not simply be named Table, conflicts with DocumentFormat.OpenXml.Drawing.Table.</remarks>
    public class PptxTable
    {
        private readonly A.Table tbl;

        public PptxTable(A.Table tbl)
        {
            this.tbl = tbl;
        }

        public class Cell
        {
            public string Tag { get; set; }

            public string NewText { get; set; }

            public Cell(string tag, string newText)
            {
                this.Tag = tag;
                this.NewText = newText;
            }
        }

        public void AppendRow(Cell[] cells)
        {
            A.TableRow tr = (A.TableRow)this.GetSecondRow().CloneNode(true);

            foreach (A.Paragraph p in tr.Descendants<A.Paragraph>())
            {
                foreach (var cell in cells)
                {
                    PptxSlide.ReplaceTagInParagraph(p, cell.Tag, cell.NewText);
                }
            }

            this.tbl.AppendChild(tr);
        }

        private A.TableRow GetSecondRow()
        {
            // TODO check for error and throw an nice exception saying the template table
            // is erronous and does not contains two rows (titles + first row)
            A.TableRow tr = this.tbl.Descendants<A.TableRow>().ElementAt(1);
            return tr;
        }
    }
}

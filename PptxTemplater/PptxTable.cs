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

        public void AppendRow(params string[] cells)
        {
            A.TableRow tr = (A.TableRow)this.GetSecondRow().CloneNode(false);
            foreach (var cell in cells)
            {
                tr.AppendChild(CreateTextCell(cell));
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

        /// <summary>
        /// Creates a table cell that contains a text.
        /// </summary>
        /// <see href="http://blogs.msdn.com/b/brian_jones/archive/2009/08/13/adding-repeating-data-to-powerpoint.aspx">Adding Repeating Data to PowerPoint</see>
        private static A.TableCell CreateTextCell(string text)
        {
            A.TableCell tc =
                new A.TableCell(
                    new A.TextBody(
                        new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text(text)))),
                    new A.TableCellProperties());
            return tc;
        }
    }
}

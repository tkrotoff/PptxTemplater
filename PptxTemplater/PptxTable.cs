namespace PptxTemplater
{
    using System.Collections.Generic;
    using System.Linq;

    using A = DocumentFormat.OpenXml.Drawing;

    /// <summary>
    /// Represents a table inside a PowerPoint file.
    /// </summary>
    /// <remarks>
    /// Could not simply be named Table, conflicts with DocumentFormat.OpenXml.Drawing.Table.
    ///
    /// Structure of a table (3 columns x 2 lines):
    /// a:graphic
    ///  a:graphicData
    ///   a:tbl (Table)
    ///    a:tblGrid (TableGrid)
    ///     a:gridCol (GridColumn)
    ///     a:gridCol
    ///     a:gridCol
    ///    a:tr (TableRow)
    ///     a:tc (TableCell)
    ///     a:tc
    ///     a:tc
    ///    a:tr
    ///     a:tc
    ///     a:tc
    ///     a:tc
    ///
    /// <![CDATA[
    /// <a:graphic>
    ///   <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
    ///     <a:tbl>
    ///       <a:tblPr firstRow="1" bandRow="1">
    ///         <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>
    ///       </a:tblPr>
    ///       <a:tblGrid>
    ///         <a:gridCol w="2743200" />
    ///         <a:gridCol w="2743200" />
    ///         <a:gridCol w="2743200" />
    ///       </a:tblGrid>
    ///       <a:tr h="370840">
    ///         <a:tc>
    ///           <a:txBody>
    ///             <a:bodyPr />
    ///             <a:lstStyle />
    ///             [PptxParagraph]
    ///           </a:txBody>
    ///           <a:tcPr />
    ///         </a:tc>
    ///         <a:tc>
    ///           <a:txBody>
    ///             <a:bodyPr />
    ///             <a:lstStyle />
    ///             [PptxParagraph]
    ///           </a:txBody>
    ///           <a:tcPr />
    ///         </a:tc>
    ///         <a:tc>
    ///           <a:txBody>
    ///             <a:bodyPr />
    ///             <a:lstStyle />
    ///             [PptxParagraph]
    ///           </a:txBody>
    ///           <a:tcPr />
    ///         </a:tc>
    ///       </a:tr>
    ///       <a:tr h="370840">
    ///         <a:tc>
    ///           <a:txBody>
    ///             <a:bodyPr />
    ///             <a:lstStyle />
    ///             [PptxParagraph]
    ///           </a:txBody>
    ///           <a:tcPr />
    ///         </a:tc>
    ///         <a:tc>
    ///           <a:txBody>
    ///             <a:bodyPr />
    ///             <a:lstStyle />
    ///             [PptxParagraph]
    ///           </a:txBody>
    ///           <a:tcPr />
    ///         </a:tc>
    ///         <a:tc>
    ///           <a:txBody>
    ///             <a:bodyPr />
    ///             <a:lstStyle />
    ///             [PptxParagraph]
    ///           </a:txBody>
    ///           <a:tcPr />
    ///         </a:tc>
    ///       </a:tr>
    ///     </a:tbl>
    ///   </a:graphicData>
    /// </a:graphic>
    /// ]]>
    /// </remarks>
    public class PptxTable
    {
        private readonly PptxSlide slideTemplate;
        private readonly int tblId;

        public string Title { get; private set; }

        internal PptxTable(PptxSlide slideTemplate, int tblId, string title)
        {
            this.slideTemplate = slideTemplate;
            this.tblId = tblId;
            this.Title = title;
        }

        /// <summary>
        /// Represents a cell inside a table.
        /// </summary>
        public class Cell
        {
            internal string Tag { get; private set; }

            internal string NewText { get; private set; }

            public Cell(string tag, string newText)
            {
                this.Tag = tag;
                this.NewText = newText;
            }
        }

        /// <summary>
        /// Removes the given columns.
        /// </summary>
        /// <param name="columns">Indexes of the columns to remove.</param>
        public void RemoveColumns(IEnumerable<int> columns)
        {
            A.Table tbl = this.slideTemplate.FindTable(this.tblId);
            A.TableGrid tblGrid = tbl.TableGrid;

            // Remove the latest columns first
            IEnumerable<int> columnsSorted = from column in columns
                                             orderby column descending
                                             select column;

            foreach (int column in columnsSorted)
            {
                for (int row = 0; row < RowsCount(tbl); row++)
                {
                    A.TableRow tr = GetRow(tbl, row);

                    // Remove the column from the row
                    A.TableCell tc = GetCell(tr, column);
                    tc.Remove();
                }

                // Remove the column from TableGrid
                A.GridColumn gridCol = tblGrid.Descendants<A.GridColumn>().ElementAt(column);
                gridCol.Remove();
            }

            this.slideTemplate.Save();
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
            for (int i = 0; i < rows.Count();)
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

        /// <summary>
        /// Gets the columns titles as an array of strings.
        /// </summary>
        public string[] ColumnTitles()
        {
            List<string> titles = new List<string>();

            A.Table tbl = this.slideTemplate.FindTable(this.tblId);
            A.TableRow tr = GetRow(tbl, 0); // The first table row == the columns
            foreach (A.Paragraph p in tr.Descendants<A.Paragraph>())
            {
                string columnTitle = PptxParagraph.GetTexts(p);
                titles.Add(columnTitle);
            }

            return titles.ToArray();
        }

        /// <summary>
        /// Helper method.
        /// </summary>
        private static int RowsCount(A.Table tbl)
        {
            return tbl.Descendants<A.TableRow>().Count();
        }

        /// <summary>
        /// Helper method.
        /// </summary>
        private static A.TableRow GetRow(A.Table tbl, int row)
        {
            A.TableRow tr = tbl.Descendants<A.TableRow>().ElementAt(row);
            return tr;
        }

        /// <summary>
        /// Helper method.
        /// </summary>
        private static A.TableCell GetCell(A.TableRow tr, int column)
        {
            A.TableCell tc = tr.Descendants<A.TableCell>().ElementAt(column);
            return tc;
        }
    }
}

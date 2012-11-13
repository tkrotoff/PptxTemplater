namespace PptxTemplater
{
    using System.Collections.Generic;
    using System.Linq;

    using DocumentFormat.OpenXml.Packaging;

    using A = DocumentFormat.OpenXml.Drawing;

    /// <summary>
    /// Represents a table inside a PowerPoint file.
    /// </summary>
    /// <remarks>
    /// Could not simply be named Table, conflicts with DocumentFormat.OpenXml.Drawing.Table.
    ///
    /// Structure of a table (3 columns x 2 lines):
    /// p:graphicFrame
    ///  a:graphic
    ///   a:graphicData
    ///    a:tbl (Table)
    ///     a:tblGrid (TableGrid)
    ///      a:gridCol (GridColumn)
    ///      a:gridCol
    ///      a:gridCol
    ///     a:tr (TableRow)
    ///      a:tc (TableCell)
    ///      a:tc
    ///      a:tc
    ///     a:tr
    ///      a:tc
    ///      a:tc
    ///      a:tc
    ///
    /// <![CDATA[
    /// <p:graphicFrame>
    ///   <p:nvGraphicFramePr>
    ///     <p:cNvPr id="22" name="Tableau 13" title="TableTitle"/>
    ///     <p:cNvGraphicFramePr>
    ///       <a:graphicFrameLocks noGrp="1"/>
    ///     </p:cNvGraphicFramePr>
    ///     <p:nvPr>
    ///       <p:extLst>
    ///       <p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}">
    ///         <p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1057920305"/>
    ///       </p:ext>
    ///       </p:extLst>
    ///     </p:nvPr>
    ///   </p:nvGraphicFramePr>
    ///   <p:xfrm>
    ///     <a:off x="4139952" y="836712"/>
    ///     <a:ext cx="4621104" cy="1538368"/>
    ///   </p:xfrm>
    ///   <a:graphic>
    ///     <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
    ///       <a:tbl>
    ///         <a:tblPr firstRow="1" bandRow="1">
    ///           <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>
    ///         </a:tblPr>
    ///         <a:tblGrid>
    ///           <a:gridCol w="2743200" />
    ///           <a:gridCol w="2743200" />
    ///           <a:gridCol w="2743200" />
    ///         </a:tblGrid>
    ///         <a:tr h="370840">
    ///           <a:tc>
    ///             <a:txBody>
    ///               <a:bodyPr />
    ///               <a:lstStyle />
    ///               [PptxParagraph]
    ///             </a:txBody>
    ///             <a:tcPr />
    ///           </a:tc>
    ///           <a:tc>
    ///             <a:txBody>
    ///               <a:bodyPr />
    ///               <a:lstStyle />
    ///               [PptxParagraph]
    ///             </a:txBody>
    ///             <a:tcPr />
    ///           </a:tc>
    ///           <a:tc>
    ///             <a:txBody>
    ///               <a:bodyPr />
    ///               <a:lstStyle />
    ///               [PptxParagraph]
    ///             </a:txBody>
    ///             <a:tcPr />
    ///           </a:tc>
    ///         </a:tr>
    ///         <a:tr h="370840">
    ///           <a:tc>
    ///             <a:txBody>
    ///               <a:bodyPr />
    ///               <a:lstStyle />
    ///               [PptxParagraph]
    ///             </a:txBody>
    ///             <a:tcPr />
    ///           </a:tc>
    ///           <a:tc>
    ///             <a:txBody>
    ///               <a:bodyPr />
    ///               <a:lstStyle />
    ///               [PptxParagraph]
    ///             </a:txBody>
    ///             <a:tcPr />
    ///           </a:tc>
    ///           <a:tc>
    ///             <a:txBody>
    ///               <a:bodyPr />
    ///               <a:lstStyle />
    ///               [PptxParagraph]
    ///             </a:txBody>
    ///             <a:tcPr />
    ///           </a:tc>
    ///         </a:tr>
    ///       </a:tbl>
    ///     </a:graphicData>
    ///   </a:graphic>
    /// </p:graphicFrame>
    /// ]]>
    /// </remarks>
    public class PptxTable
    {
        private PptxSlide slideTemplate;

        private readonly int tblId;

        public string Title { get; private set; }

        internal PptxTable(PptxSlide slideTemplate, int tblId, string title)
        {
            this.slideTemplate = slideTemplate;
            this.tblId = tblId;
            this.Title = title;
        }

        /// <summary>
        /// Represents a cell inside a table (a:tbl).
        /// </summary>
        public class Cell
        {
            internal string Tag { get; private set; }

            internal string NewText { get; private set; }

            public class BackgroundPicture
            {
                public byte[] Picture { get; set; }
                public string ContentType { get; set; }
                public int Top { get; set; }
                public int Right { get; set; }
                public int Bottom { get; set; }
                public int Left { get; set; }
            }

            internal BackgroundPicture Picture { get; private set; }

            public Cell(string tag, string newText)
            {
                this.Tag = tag;
                this.NewText = newText;
            }

            public Cell(string tag, string newText, BackgroundPicture backgroundPicture)
            {
                this.Tag = tag;
                this.NewText = newText;
                this.Picture = backgroundPicture;
            }
        }

        /// <summary>
        /// Removes the table from the slide.
        /// </summary>
        /// <returns>True if the table has been removed; false otherwise.</returns>
        public bool Remove()
        {
            return this.slideTemplate.RemoveTable(this.tblId);
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
        /// Sets a background picture for a table cell (a:tc).
        /// </summary>
        /// <remarks>
        /// <![CDATA[
        /// <a:tc>
        ///  <a:txBody>
        ///   <a:bodyPr/>
        ///   <a:lstStyle/>
        ///   <a:p>
        ///    <a:endParaRPr lang="fr-FR" dirty="0"/>
        ///   </a:p>
        ///  </a:txBody>
        ///  <a:tcPr> (TableCellProperties)
        ///   <a:blipFill dpi="0" rotWithShape="1">
        ///    <a:blip r:embed="rId2"/>
        ///    <a:srcRect/>
        ///    <a:stretch>
        ///     <a:fillRect b="12000" r="90000" t="14000"/>
        ///    </a:stretch>
        ///   </a:blipFill>
        ///  </a:tcPr>
        /// </a:tc>
        /// ]]>
        /// </remarks>
        private static void SetTableCellPropertiesWithBackgroundPicture(PptxSlide slide, A.TableCellProperties tcPr, Cell.BackgroundPicture backgroundPicture)
        {
            if (backgroundPicture.Picture == null)
            {
                return;
            }

            ImagePart imagePart = slide.AddPicture(backgroundPicture.Picture, backgroundPicture.ContentType);

            A.BlipFill blipFill = new A.BlipFill();
            A.Blip blip = new A.Blip() { Embed = slide.GetIdOfImagePart(imagePart) };
            A.SourceRectangle srcRect = new A.SourceRectangle();
            A.Stretch stretch = new A.Stretch();
            A.FillRectangle fillRect = new A.FillRectangle()
                {
                    Top = backgroundPicture.Top,
                    Right = backgroundPicture.Right,
                    Bottom = backgroundPicture.Bottom,
                    Left = backgroundPicture.Left
                };
            stretch.AppendChild(fillRect);
            blipFill.AppendChild(blip);
            blipFill.AppendChild(srcRect);
            blipFill.AppendChild(stretch);
            tcPr.AppendChild(blipFill);
        }

        /// <summary>
        /// Replaces a tag inside the table (a:tbl).
        /// </summary>
        /// <param name="cell">Contains the tag, the new text and a pciture.</param>
        /// <returns>True if a tag has been found and replaced; false otherwise.</returns>
        public bool ReplaceTag(Cell cell)
        {
            bool replacedAtLeastOnce = false;

            PptxSlide slide = this.slideTemplate;
            A.Table tbl = slide.FindTable(this.tblId);

            // a:tr
            foreach (A.TableRow tr in tbl.Descendants<A.TableRow>())
            {
                // a:tc
                foreach (A.TableCell tc in tr.Descendants<A.TableCell>())
                {
                    bool replaced = ReplaceTag(slide, tc, cell);
                    if (replaced)
                    {
                        replacedAtLeastOnce = true;
                    }
                }
            }

            return replacedAtLeastOnce;
        }

        /// <summary>
        /// Replaces a tag inside a given table cell (a:tc).
        /// </summary>
        /// <param name="slide">The PptxSlide.</param>
        /// <param name="tc">The table cell (a:tc).</param>
        /// <param name="cell">Contains the tag, the new text and a picture.</param>
        /// <returns>True if a tag has been found and replaced; false otherwise.</returns>
        private static bool ReplaceTag(PptxSlide slide, A.TableCell tc, Cell cell)
        {
            bool replacedAtLeastOnce = false;

            // a:p
            foreach (A.Paragraph p in tc.Descendants<A.Paragraph>())
            {
                bool replaced = PptxParagraph.ReplaceTag(p, cell.Tag, cell.NewText);
                if (replaced)
                {
                    replacedAtLeastOnce = true;

                    // a:tcPr
                    if (cell.Picture != null)
                    {
                        A.TableCellProperties tcPr = tc.GetFirstChild<A.TableCellProperties>();
                        SetTableCellPropertiesWithBackgroundPicture(slide, tcPr, cell.Picture);
                    }
                }
            }

            return replacedAtLeastOnce;
        }

        /// <summary>
        /// Changes the cells from the table (tbl).
        /// </summary>
        /// <returns>The list of remaining rows that could not be inserted, you will have to create a new slide.</returns>
        public List<Cell[]> SetRowsNoInsert(IList<Cell[]> rows)
        {
            PptxSlide slide = this.slideTemplate;
            A.Table tbl = slide.FindTable(this.tblId);

            // done starts at 1 instead of 0 because we don't care about the first row
            // The first row contains the titles for the columns
            int done = 1;
            for (int i = 0; i < rows.Count(); i++)
            {
                Cell[] row = rows[i];

                if (done < RowsCount(tbl))
                {
                    // a:tr
                    A.TableRow tr = GetRow(tbl, done);

                    // a:tc
                    foreach (A.TableCell tc in tr.Descendants<A.TableCell>())
                    {
                        foreach (Cell cell in row)
                        {
                            ReplaceTag(slide, tc, cell);
                        }
                    }

                    done++;
                }
                else
                {
                    break;
                }
            }

            // Remove the last remaining rows if any
            for (int row = RowsCount(tbl) - 1; row >= done; row--)
            {
                A.TableRow tr = GetRow(tbl, row);
                tr.Remove();
            }

            // Save the latest slide
            // Mandatory otherwise the next time SetRows() is run (on a different table)
            // the rows from the previous tables will not contained the right data (from PptxParagraph.ReplaceTag())
            slide.Save();

            // Computes the remaining rows if any
            List<Cell[]> remainingRows = new List<Cell[]>();
            for (int row = done - 1; row < rows.Count; row++)
            {
                remainingRows.Add(rows[row]);
            }

            return remainingRows;
        }

        /// <summary>
        /// Changes the cells from the table (tbl).
        /// </summary>
        /// <remarks>
        /// Be careful when calling this method multiple times.
        /// This method can potentially change the number of slides (by inserting new slides) so you are better off
        /// calling it last.
        /// </remarks>
        /// <returns>The list of inserted (new) slides.</returns>
        public List<PptxSlide> SetRows(IList<Cell[]> rows)
        {
            List<PptxSlide> insertedSlides = new List<PptxSlide>();

            // Create a new slide from the template slide
            PptxSlide slide = this.slideTemplate.Clone();
            insertedSlides.Add(slide);
            PptxSlide.InsertAfter(slide, this.slideTemplate);
            A.Table tbl = slide.FindTable(this.tblId);

            // donePerSlide starts at 1 instead of 0 because we don't care about the first row
            // The first row contains the titles for the columns
            int donePerSlide = 1;
            for (int i = 0; i < rows.Count();)
            {
                Cell[] row = rows[i];

                if (donePerSlide < RowsCount(tbl))
                {
                    // a:tr
                    A.TableRow tr = GetRow(tbl, donePerSlide);

                    // a:tc
                    foreach (A.TableCell tc in tr.Descendants<A.TableCell>())
                    {
                        foreach (Cell cell in row)
                        {
                            ReplaceTag(slide, tc, cell);
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
                    insertedSlides.Add(newSlide);
                    PptxSlide.InsertAfter(newSlide, slide);
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

            // Remove the template slide
            this.slideTemplate.Remove();

            return insertedSlides;
        }

        /// <summary>
        /// Gets the columns titles as an array of strings.
        /// </summary>
        public IEnumerable<string> ColumnTitles()
        {
            List<string> titles = new List<string>();

            A.Table tbl = this.slideTemplate.FindTable(this.tblId);
            A.TableRow tr = GetRow(tbl, 0); // The first table row == the columns
            foreach (A.Paragraph p in tr.Descendants<A.Paragraph>())
            {
                string columnTitle = PptxParagraph.GetTexts(p);
                titles.Add(columnTitle);
            }

            return titles;
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

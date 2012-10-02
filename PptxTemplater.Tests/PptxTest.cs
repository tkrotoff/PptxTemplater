namespace PptxTemplater.Tests
{
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Text;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class PptxTest
    {
        private void AssertPptxEquals(string file, int nbSlides, string expected)
        {
            Pptx pptx = new Pptx(file, false);
            Assert.AreEqual(nbSlides, pptx.SlidesCount());
            StringBuilder result = new StringBuilder();
            for (int i = 0; i < nbSlides; i++)
            {
                string[] texts = pptx.GetTextsInSlide(i);
                result.Append(string.Join(" ", texts));
                result.Append(" ");
            }
            pptx.Close();
            Assert.AreEqual(expected, result.ToString());
        }

        [TestMethod]
        [ExpectedException(typeof(FileFormatException), "File contains corrupted data.")]
        public void FileFormatException()
        {
            Pptx pptx = new Pptx("../../files/picture1.png", false);
            pptx.Close();
        }

        [TestMethod]
        public void EmptyPowerPoint()
        {
            const string file = "../../files/EmptyPowerPoint.pptx";
            const string thumbnail_empty_png = "../../files/thumbnail_empty.png";
            const string thumbnail_empty_output_png = "../../files/thumbnail_empty_output.png";

            Pptx pptx = new Pptx(file, false);
            int nbSlides = pptx.SlidesCount();
            Assert.AreEqual(0, nbSlides);

            byte[] thumbnail_empty_output = pptx.GetThumbnail();
            File.WriteAllBytes(thumbnail_empty_output_png, thumbnail_empty_output);
            byte[] thumbnail_empty = File.ReadAllBytes(thumbnail_empty_png);
            CollectionAssert.AreEqual(thumbnail_empty, thumbnail_empty_output);

            pptx.Close();
        }

        [TestMethod]
        public void GetTextsInAllSlides()
        {
            const string file = "../../files/GetTextsInAllSlides.pptx";

            Pptx pptx = new Pptx(file, false);
            int nbSlides = pptx.SlidesCount();
            Assert.AreEqual(3, nbSlides);

            var slidesTexts = new Dictionary<int, string[]>();
            for (int i = 0; i < nbSlides; i++)
            {
                string[] texts = pptx.GetTextsInSlide(i);
                slidesTexts.Add(i, texts);
            }

            string[] expected = { "test1", "Hello, world!" };
            CollectionAssert.AreEqual(expected, slidesTexts[0]);

            expected = new string[]
                           {
                               "Title 1", "Bullet 1", "Bullet 2",
                               "Column 1", "Column 2", "Column 3", "Column 4", "Column 5",
                               "Line 1", "", "", "", "",
                               "Line 2", "", "", "", "",
                               "Line 3", "", "", "", "",
                               "Line 4", "", "", "", ""
                           };
            CollectionAssert.AreEqual(expected, slidesTexts[1]);

            expected = new string[] { "Title 2", "Bullet 1", "Bullet 2", "Comment ça va ?" };
            CollectionAssert.AreEqual(expected, slidesTexts[2]);

            pptx.Close();
        }

        [TestMethod]
        public void GetNotesInAllSlides()
        {
            const string file = "../../files/GetNotesInAllSlides.pptx";

            Pptx pptx = new Pptx(file, false);
            int nbSlides = pptx.SlidesCount();
            Assert.AreEqual(4, nbSlides);

            var slidesNotes = new Dictionary<int, string[]>();
            for (int i = 0; i < nbSlides; i++)
            {
                string[] notes = pptx.GetNotesInSlide(i);
                slidesNotes.Add(i, notes);
            }

            string[] expected = { "Bonjour", "{{comment1}}", "Hello", "1" };
            CollectionAssert.AreEqual(expected, slidesNotes[0]);

            expected = new string[] { "{{comment2}}", "2" };
            CollectionAssert.AreEqual(expected, slidesNotes[1]);

            expected = new string[] { };
            CollectionAssert.AreEqual(expected, slidesNotes[2]);

            // TODO Why "Comment çava ?" instead of "Comment ça va ?"
            expected = new string[] { "Bonjour {{comment3}} Hello", "Comment çava ?", "", "", "Hola!", "", "4" };
            CollectionAssert.AreEqual(expected, slidesNotes[3]);

            pptx.Close();
        }

        [TestMethod]
        public void GetTablesInAllSlides()
        {
            const string file = "../../files/ReplaceTablesInAllSlides.pptx";

            Pptx pptx = new Pptx(file, false);
            int nbSlides = pptx.SlidesCount();
            Assert.AreEqual(3, nbSlides);

            var slidesTables = new Dictionary<int, PptxTable[]>();
            for (int i = 0; i < nbSlides; i++)
            {
                PptxTable[] tables = pptx.GetTablesInSlide(i);
                slidesTables.Add(i, tables);
            }

            string[] expected = { "Table1", "Col2", "Col3", "Col4", "Col5", "Col6" };
            CollectionAssert.AreEqual(expected, slidesTables[0][0].ColumnTitles());

            expected = new string[] { "Table2", "Col2", "Col3", "Col4", "Col5", "Col6" };
            CollectionAssert.AreEqual(expected, slidesTables[1][0].ColumnTitles());

            expected = new string[] { "Table3", "Col2", "Col3", "Col4", "Col5", "Col6" };
            CollectionAssert.AreEqual(expected, slidesTables[1][1].ColumnTitles());

            pptx.Close();
        }

        [TestMethod]
        public void ReplaceTagsInAllSlides()
        {
            const string srcFileName = "../../files/ReplaceTagsInAllSlides.pptx";
            const string dstFileName = "../../files/ReplaceTagsInAllSlides_output.pptx";
            File.Delete(dstFileName);
            File.Copy(srcFileName, dstFileName);

            Pptx pptx = new Pptx(dstFileName, true);
            int nbSlides = pptx.SlidesCount();
            Assert.AreEqual(3, nbSlides);

            // First slide
            pptx.ReplaceTagInSlide(0, "{{hello}}", "HELLO HOW ARE YOU?");
            pptx.ReplaceTagInSlide(0, "{{bonjour}}", "BONJOUR TOUT LE MONDE");
            pptx.ReplaceTagInSlide(0, "{{hola}}", "HOLA MAMA QUE TAL?");

            // Second slide
            pptx.ReplaceTagInSlide(1, "{{hello}}", "H");
            pptx.ReplaceTagInSlide(1, "{{bonjour}}", "B");
            pptx.ReplaceTagInSlide(1, "{{hola}}", "H");

            // Third slide
            pptx.ReplaceTagInSlide(2, "{{hello}}", string.Empty);
            pptx.ReplaceTagInSlide(2, "{{bonjour}}", string.Empty);
            pptx.ReplaceTagInSlide(2, "{{hola}}", null);
            pptx.ReplaceTagInSlide(2, null, string.Empty);
            pptx.ReplaceTagInSlide(2, null, null);

            pptx.Close();

            this.AssertPptxEquals(dstFileName, 3, "words HELLO HOW ARE YOU?|HELLO HOW ARE YOU?|HOLA MAMA QUE TAL?, world! A tag {{hoHOLA MAMA QUE TAL?la}} inside a sentence BONJOUR TOUT LE MONDE A tag BONJOUR TOUT LE MONDEHOLA MAMA QUE TAL?BONJOUR TOUT LE MONDE inside a sentence HELLO HOW ARE YOU?, world! words H|H|H, world! A tag {{hoHla}} inside a sentence B A tag BHB inside a sentence H, world! words ||, world! A tag  inside a sentence  A tag inside a sentence , world! ");
        }

        [TestMethod]
        public void ReplacePicturesInAllSlides()
        {
            const string srcFileName = "../../files/ReplacePicturesInAllSlides.pptx";
            const string dstFileName = "../../files/ReplacePicturesInAllSlides_output.pptx";
            File.Delete(dstFileName);
            File.Copy(srcFileName, dstFileName);

            Pptx pptx = new Pptx(dstFileName, true);
            int nbSlides = pptx.SlidesCount();
            Assert.AreEqual(2, nbSlides);

            pptx.ReplaceTagInSlide(0, "{{hello}}", "HELLO!");

            const string picture1_replace_png = "../../files/picture1_replace.png";
            const string picture1_replace_png_contentType = "image/png";
            const string picture1_replace_bmp = "../../files/picture1_replace.bmp";
            const string picture1_replace_bmp_contentType = "image/bmp";
            const string picture1_replace_jpeg = "../../files/picture1_replace.jpeg";
            const string picture1_replace_jpeg_contentType = "image/jpeg";
            const Stream picture1_replace_null = null;
            for (int i = 0; i < nbSlides; i++)
            {
                pptx.ReplacePictureInSlide(i, "{{picture1png}}", picture1_replace_png, picture1_replace_png_contentType);
                pptx.ReplacePictureInSlide(i, "{{picture1bmp}}", picture1_replace_bmp, picture1_replace_bmp_contentType);
                pptx.ReplacePictureInSlide(i, "{{picture1jpeg}}", picture1_replace_jpeg, picture1_replace_jpeg_contentType);

                pptx.ReplacePictureInSlide(i, null, picture1_replace_png, picture1_replace_png_contentType);
                pptx.ReplacePictureInSlide(i, "{{picture1null}}", picture1_replace_null, picture1_replace_png_contentType);
                pptx.ReplacePictureInSlide(i, "{{picture1null}}", picture1_replace_png, null);
                pptx.ReplacePictureInSlide(i, "{{picture1null}}", picture1_replace_null, null);
            }

            pptx.Close();

            // Sorry, you will have to manually check that the pictures have been replaced
        }

        [TestMethod]
        public void RemoveColumns()
        {
            const string srcFileName = "../../files/RemoveColumns.pptx";
            const string dstFileName = "../../files/RemoveColumns_output.pptx";
            File.Delete(dstFileName);
            File.Copy(srcFileName, dstFileName);

            Pptx pptx = new Pptx(dstFileName, true);

            PptxTable[] tables = pptx.FindTables("{{table1}}");
            foreach (PptxTable table in tables)
            {
                int[] columns = new int[] { 1, 3 };
                table.RemoveColumns(columns);
            }

            pptx.Close();

            this.AssertPptxEquals(dstFileName, 1, "Column 0 Column2 Column 4 Cell 1.0 Cell 1.2 Cell 1.4 Cell 2.0 Cell 2.2 Cell 2.4 Cell 3.0 Cell 3.2 Cell 3.4 Cell 4.0 Cell 4.2 Cell 4.4 Cell 5.0 Cell 5.2 Cell 5.4 ");
        }

        [TestMethod]
        public void ReplaceTablesInAllSlides()
        {
            const string srcFileName = "../../files/ReplaceTablesInAllSlides.pptx";
            const string dstFileName = "../../files/ReplaceTablesInAllSlides_output.pptx";
            File.Delete(dstFileName);
            File.Copy(srcFileName, dstFileName);

            Pptx pptx = new Pptx(dstFileName, true);

            // Change the tags before to insert rows
            // PptxTable.SetRows() might change the number of slides inside the presentation
            pptx.ReplaceTagInSlide(0, "{{hello}}", "HELLO!");

            // Change the pictures before to insert rows
            // PptxTable.SetRows() might change the number of slides inside the presentation
            const string picture1_replace_png = "../../files/picture1_replace.png";
            const string picture1_replace_png_contentType = "image/png";
            pptx.ReplacePictureInSlide(2, "{{picture1png}}", picture1_replace_png, picture1_replace_png_contentType);

            PptxTable[] tables;
            List<PptxTable.Cell[]> rows;
            PptxTable.Cell[] row;

            tables = pptx.FindTables("{{table1}}");
            rows = new List<PptxTable.Cell[]>();
            row = new[]
                {
                    new PptxTable.Cell("{{cell1}}", "Hello, world! 1"),
                    new PptxTable.Cell("{{cell2}}", "Hello, world! 2"),
                    new PptxTable.Cell("{{cell3}}", "Hello, world! 3"),
                    new PptxTable.Cell("{{cell4}}", "Hello, world! 4"),
                    new PptxTable.Cell("{{cell5}}", "Hello, world! 5"),
                    new PptxTable.Cell("{{cell6}}", "Hello, world! 6"),

                    new PptxTable.Cell(null, "null"),
                    new PptxTable.Cell("{{unknown}}", null),
                    new PptxTable.Cell(null, null)
                };
            rows.Add(row);
            rows.Add(row);
            rows.Add(row);
            rows.Add(row);
            rows.Add(row);
            rows.Add(row);
            rows.Add(row);
            rows.Add(row);
            rows.Add(row);
            rows.Add(row);
            foreach (PptxTable table in tables)
            {
                table.SetRows(rows);
            }

            tables = pptx.FindTables("{{table2}}");
            rows = new List<PptxTable.Cell[]>();
            row = new[]
                {
                    new PptxTable.Cell("{{cell1}}", "Bonjour 1"),
                    new PptxTable.Cell("{{cell2}}", "Bonjour 2"),
                    new PptxTable.Cell("{{cell3}}", "Bonjour 3"),
                    new PptxTable.Cell("{{cell4}}", "Bonjour 4"),
                    new PptxTable.Cell("{{cell5}}", "Bonjour 5"),
                    new PptxTable.Cell("{{cell6}}", "Bonjour 6")
                };
            rows.Add(row);
            rows.Add(row);
            foreach (PptxTable table in tables)
            {
                table.SetRows(rows);
            }

            tables = pptx.FindTables("{{table3}}");
            rows = new List<PptxTable.Cell[]>();
            row = new[]
                {
                    new PptxTable.Cell("{{cell1}}", "Hola! 1"),
                    new PptxTable.Cell("{{cell2}}", "Hola! 2"),
                    new PptxTable.Cell("{{cell3}}", "Hola! 3"),
                    new PptxTable.Cell("{{cell4}}", "Hola! 4"),
                    new PptxTable.Cell("{{cell5}}", "Hola! 5"),
                    new PptxTable.Cell("{{cell6}}", "Hola! 6")
                };
            rows.Add(row);
            rows.Add(row);
            rows.Add(row);
            rows.Add(row);
            rows.Add(row);
            rows.Add(row);
            rows.Add(row);
            rows.Add(row);
            rows.Add(row);
            rows.Add(row);
            foreach (PptxTable table in tables)
            {
                table.SetRows(rows);
            }

            pptx.Close();

            this.AssertPptxEquals(dstFileName, 7, "Table1 Col2 Col3 Col4 Col5 Col6 HELLO Hello, world! 1  Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 HELLO Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 HELLO! Table1 Col2 Col3 Col4 Col5 Col6 HELLO Hello, world! 1  Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 HELLO Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 HELLO! Table1 Col2 Col3 Col4 Col5 Col6 HELLO Hello, world! 1  Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 HELLO! Table2 Col2 Col3 Col4 Col5 Col6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Table3 Col2 Col3 Col4 Col5 Col6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Table2 Col2 Col3 Col4 Col5 Col6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Table3 Col2 Col3 Col4 Col5 Col6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Table2 Col2 Col3 Col4 Col5 Col6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Table3 Col2 Col3 Col4 Col5 Col6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6  ");
        }

        [TestMethod]
        public void GetThumbnail()
        {
            string file = "../../files/GetTextsInAllSlides.pptx";
            const string thumbnail_default_png = "../../files/thumbnail_default.png";
            const string thumbnail_default_output_png = "../../files/thumbnail_default_output.png";
            const string thumbnail_128x96_png = "../../files/thumbnail_128x96.png";
            const string thumbnail_128x96_output_png = "../../files/thumbnail_128x96_output.png";
            const string thumbnail_512x384_png = "../../files/thumbnail_512x384.png";
            const string thumbnail_512x384_output_png = "../../files/thumbnail_512x384_output.png";

            Pptx pptx = new Pptx(file, false);
            byte[] thumbnail_default_output = pptx.GetThumbnail(); // Default size
            File.WriteAllBytes(thumbnail_default_output_png, thumbnail_default_output);
            byte[] thumbnail_128x96_output = pptx.GetThumbnail(new Size(128, 96));
            File.WriteAllBytes(thumbnail_128x96_output_png, thumbnail_128x96_output);
            byte[] thumbnail_512x384_output = pptx.GetThumbnail(new Size(512, 384));
            File.WriteAllBytes(thumbnail_512x384_output_png, thumbnail_512x384_output);

            // Check the generated thumbnail are ok
            byte[] thumbnail_default = File.ReadAllBytes(thumbnail_default_png);
            CollectionAssert.AreEqual(thumbnail_default, thumbnail_default_output);
            byte[] thumbnail_128x96 = File.ReadAllBytes(thumbnail_128x96_png);
            CollectionAssert.AreEqual(thumbnail_128x96, thumbnail_128x96_output);
            byte[] thumbnail_512x384 = File.ReadAllBytes(thumbnail_512x384_png); // Will look blurry
            CollectionAssert.AreEqual(thumbnail_512x384, thumbnail_512x384_output);

            pptx.Close();

            // Test a 16/10 portrait PowerPoint file
            file = "../../files/portrait_16_10.pptx";
            const string thumbnail_portrait_16_10_png = "../../files/thumbnail_portrait_16_10.png";
            const string thumbnail_portrait_16_10_output_png = "../../files/thumbnail_portrait_16_10_output.png";

            pptx = new Pptx(file, false);
            byte[] thumbnail_portrait_16_10_output = pptx.GetThumbnail(); // Default size
            File.WriteAllBytes(thumbnail_portrait_16_10_output_png, thumbnail_portrait_16_10_output);

            byte[] thumbnail_portrait_16_10 = File.ReadAllBytes(thumbnail_portrait_16_10_png);
            CollectionAssert.AreEqual(thumbnail_portrait_16_10, thumbnail_portrait_16_10_output);

            pptx.Close();
        }
    }
}

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
        public void GetAllTextInAllSlides()
        {
            const string file = "../../files/GetAllTextInAllSlides.pptx";

            Pptx pptx = new Pptx(file, false);
            int nbSlides = pptx.SlidesCount();
            Assert.AreEqual(3, nbSlides);

            var slidesTexts = new Dictionary<int, string[]>();
            for (int i = 0; i < nbSlides; i++)
            {
                string[] texts = pptx.GetAllTextInSlide(i);
                slidesTexts.Add(i, texts);
            }

            string[] expected = {"test1", "Hello, world!"};
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
            expected = new string[] {"Title 2", "Bullet 1", "Bullet 2"};
            CollectionAssert.AreEqual(expected, slidesTexts[2]);

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
            pptx.ReplaceTagInSlide(2, "{{hola}}", string.Empty);

            pptx.Close();

            // Check the replaced texts are here
            pptx = new Pptx(dstFileName, false);
            nbSlides = pptx.SlidesCount();
            Assert.AreEqual(3, nbSlides);
            StringBuilder result = new StringBuilder();
            for (int i = 0; i < nbSlides; i++)
            {
                string[] texts = pptx.GetAllTextInSlide(i);
                result.Append(string.Join(" ", texts));
                result.Append(" ");
            }
            pptx.Close();
            const string expected = "words HELLO HOW ARE YOU?|HELLO HOW ARE YOU?|HOLA MAMA QUE TAL?, world! A tag {{hoHOLA MAMA QUE TAL?la}} inside a sentence BONJOUR TOUT LE MONDE A tag BONJOUR TOUT LE MONDEHOLA MAMA QUE TAL?BONJOUR TOUT LE MONDE inside a sentence HELLO HOW ARE YOU?, world! words H|H|H, world! A tag {{hoHla}} inside a sentence B A tag BHB inside a sentence H, world! words ||, world! A tag  inside a sentence  A tag inside a sentence , world! ";
            Assert.AreEqual(expected, result.ToString());
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

            const string picture1_replace_png = "../../files/picture1_replace.png";
            const string picture1_replace_png_contentType = "image/png";
            const string picture1_replace_bmp = "../../files/picture1_replace.bmp";
            const string picture1_replace_bmp_contentType = "image/bmp";
            const string picture1_replace_jpeg = "../../files/picture1_replace.jpeg";
            const string picture1_replace_jpeg_contentType = "image/jpeg";
            for (int i = 0; i < nbSlides; i++)
            {
                pptx.ReplacePictureInSlide(i, "{{picture1png}}", picture1_replace_png, picture1_replace_png_contentType);
                pptx.ReplacePictureInSlide(i, "{{picture1bmp}}", picture1_replace_bmp, picture1_replace_bmp_contentType);
                pptx.ReplacePictureInSlide(i, "{{picture1jpeg}}", picture1_replace_jpeg, picture1_replace_jpeg_contentType);
            }

            pptx.Close();

            // Sorry, you will have to manually check that the pictures have been replaced
        }

        [TestMethod]
        public void ReplaceTablesInAllSlides()
        {
            const string srcFileName = "../../files/ReplaceTablesInAllSlides.pptx";
            const string dstFileName = "../../files/ReplaceTablesInAllSlides_output.pptx";
            File.Delete(dstFileName);
            File.Copy(srcFileName, dstFileName);

            Pptx pptx = new Pptx(dstFileName, true);

            PptxTable[] tables;
            PptxTable.Cell[] row;

            tables = pptx.FindTables("{{table1}}");
            row = new[]
                {
                    new PptxTable.Cell("{{cell1}}", "Hello, world! 1"),
                    new PptxTable.Cell("{{cell2}}", "Hello, world! 2"),
                    new PptxTable.Cell("{{cell3}}", "Hello, world! 3"),
                    new PptxTable.Cell("{{cell4}}", "Hello, world! 4"),
                    new PptxTable.Cell("{{cell5}}", "Hello, world! 5"),
                    new PptxTable.Cell("{{cell6}}", "Hello, world! 6")
                };
            foreach (PptxTable table in tables)
            {
                table.SetRows(row, row, row, row, row, row, row, row, row, row);
            }

            tables = pptx.FindTables("{{table2}}");
            row = new[]
                {
                    new PptxTable.Cell("{{cell1}}", "Bonjour 1"),
                    new PptxTable.Cell("{{cell2}}", "Bonjour 2"),
                    new PptxTable.Cell("{{cell3}}", "Bonjour 3"),
                    new PptxTable.Cell("{{cell4}}", "Bonjour 4"),
                    new PptxTable.Cell("{{cell5}}", "Bonjour 5"),
                    new PptxTable.Cell("{{cell6}}", "Bonjour 6")
                };
            foreach (PptxTable table in tables)
            {
                table.SetRows(row, row);
            }

            tables = pptx.FindTables("{{table3}}");
            row = new[]
                {
                    new PptxTable.Cell("{{cell1}}", "Hola! 1"),
                    new PptxTable.Cell("{{cell2}}", "Hola! 2"),
                    new PptxTable.Cell("{{cell3}}", "Hola! 3"),
                    new PptxTable.Cell("{{cell4}}", "Hola! 4"),
                    new PptxTable.Cell("{{cell5}}", "Hola! 5"),
                    new PptxTable.Cell("{{cell6}}", "Hola! 6")
                };
            foreach (PptxTable table in tables)
            {
                table.SetRows(row, row, row, row, row, row, row, row, row, row);
            }

            pptx.Close();

            // Check the tables have been replaced
            pptx = new Pptx(dstFileName, false);
            int nbSlides = pptx.SlidesCount();
            Assert.AreEqual(6, nbSlides);
            StringBuilder result = new StringBuilder();
            for (int i = 0; i < nbSlides; i++)
            {
                string[] texts = pptx.GetAllTextInSlide(i);
                result.Append(string.Join(" ", texts));
                result.Append(" ");
            }
            pptx.Close();
            const string expected = "Table1 Col2 Col3 Col4 Col5 Col6 HELLO Hello, world! 1  Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 HELLO Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! Table1 Col2 Col3 Col4 Col5 Col6 HELLO Hello, world! 1  Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 HELLO Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! Table1 Col2 Col3 Col4 Col5 Col6 HELLO Hello, world! 1  Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! Table2 Col2 Col3 Col4 Col5 Col6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Table3 Col2 Col3 Col4 Col5 Col6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Table2 Col2 Col3 Col4 Col5 Col6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Table3 Col2 Col3 Col4 Col5 Col6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Table2 Col2 Col3 Col4 Col5 Col6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Table3 Col2 Col3 Col4 Col5 Col6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 ";
            Assert.AreEqual(expected, result.ToString());
        }

        [TestMethod]
        public void GetThumbnail()
        {
            const string file = "../../files/GetAllTextInAllSlides.pptx";
            const string thumbnail_default_png = "../../files/thumbnail_default.png";
            const string thumbnail_default_output_png = "../../files/thumbnail_default_output.png";
            const string thumbnail_128x96_png = "../../files/thumbnail_128x96.png";
            const string thumbnail_128x96_output_png = "../../files/thumbnail_128x96_output.png";

            Pptx pptx = new Pptx(file, false);
            byte[] thumbnail_default_output = pptx.GetThumbnail(); // Default size
            File.WriteAllBytes(thumbnail_default_output_png, thumbnail_default_output);
            byte[] thumbnail_128x96_output = pptx.GetThumbnail(new Size(128, 96));
            File.WriteAllBytes(thumbnail_128x96_output_png, thumbnail_128x96_output);

            // Check the generated thumbnail are ok
            byte[] thumbnail_default = File.ReadAllBytes(thumbnail_default_png);
            CollectionAssert.AreEqual(thumbnail_default, thumbnail_default_output);
            byte[] thumbnail_128x96 = File.ReadAllBytes(thumbnail_128x96_png);
            CollectionAssert.AreEqual(thumbnail_128x96, thumbnail_128x96_output);
        }
    }
}

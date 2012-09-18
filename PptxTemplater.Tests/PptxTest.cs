namespace PptxTemplater.Tests
{
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class PptxTest
    {
        [TestMethod]
        public void GetAllTextInAllSlides()
        {
            const string file = "../../files/GetAllTextInAllSlides.pptx";

            Pptx pptx = new Pptx(file, false);
            int nbSlides = pptx.CountSlides();
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
            int nbSlides = pptx.CountSlides();
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
            nbSlides = pptx.CountSlides();
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
            int nbSlides = pptx.CountSlides();
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
        public void CreateTable()
        {
            const string srcFileName = "../../files/Table5.pptx";
            const string dstFileName = "../../files/Table5_output.pptx";
            File.Delete(dstFileName);
            File.Copy(srcFileName, dstFileName);

            Pptx pptx = new Pptx(dstFileName, true);

            PptxTable[] tables = pptx.FindTables("{{table1}}");
            PptxTable.Cell[] row = new[]
                {
                    new PptxTable.Cell("{{cell1.1}}", "Hello, world! 1.1"),
                    new PptxTable.Cell("{{cell1.2}}", "Hello, world! 1.2"),
                    new PptxTable.Cell("{{cell1.3}}", "Hello, world! 1.3"),
                    new PptxTable.Cell("{{cell1.4}}", "Hello, world! 1.4"),
                    new PptxTable.Cell("{{cell1.5}}", "Hello, world! 1.5"),
                    new PptxTable.Cell("{{cell1.6}}", "Hello, world! 1.6")
                };
            foreach (PptxTable table in tables)
            {
                table.SetRows(row, row, row, row, row, row, row, row, row, row);
            }

            tables = pptx.FindTables("{{table2}}");
            row = new[]
                {
                    new PptxTable.Cell("{{cell1.1}}", "Bonjour 1.1"),
                    new PptxTable.Cell("{{cell1.2}}", "Bonjour 1.2"),
                    new PptxTable.Cell("{{cell1.3}}", "Bonjour 1.3"),
                    new PptxTable.Cell("{{cell1.4}}", "Bonjour 1.4"),
                    new PptxTable.Cell("{{cell1.5}}", "Bonjour 1.5"),
                    new PptxTable.Cell("{{cell1.6}}", "Bonjour 1.6")
                };
            foreach (PptxTable table in tables)
            {
                table.SetRows(row);
            }

            tables = pptx.FindTables("{{table3}}");
            row = new[]
                {
                    new PptxTable.Cell("{{cell1.1}}", "Hola! 1.1"),
                    new PptxTable.Cell("{{cell1.2}}", "Hola! 1.2"),
                    new PptxTable.Cell("{{cell1.3}}", "Hola! 1.3"),
                    new PptxTable.Cell("{{cell1.4}}", "Hola! 1.4"),
                    new PptxTable.Cell("{{cell1.5}}", "Hola! 1.5"),
                    new PptxTable.Cell("{{cell1.6}}", "Hola! 1.6")
                };
            foreach (PptxTable table in tables)
            {
                table.SetRows(row);
            }

            pptx.Close();
        }
    }
}

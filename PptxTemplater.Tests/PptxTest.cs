namespace PptxTemplater.Tests
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Linq;
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
                PptxSlide slide = pptx.GetSlide(i);
                string[] texts = slide.GetTexts();
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
                PptxSlide slide = pptx.GetSlide(i);
                string[] texts = slide.GetTexts();
                slidesTexts.Add(i, texts);
            }

            string[] expected = { "test1", "Hello, world!" };
            CollectionAssert.AreEqual(expected, slidesTexts[0]);

            expected = new string[]
                           {
                               "Title 1", "Bullet 1", "Bullet 2",
                               "Column 1", "Column 2", "Column 3", "Column 4", "Column 5",
                               "Line 1", string.Empty, string.Empty, string.Empty, string.Empty,
                               "Line 2", string.Empty, string.Empty, string.Empty, string.Empty,
                               "Line 3", string.Empty, string.Empty, string.Empty, string.Empty,
                               "Line 4", string.Empty, string.Empty, string.Empty, string.Empty
                           };
            CollectionAssert.AreEqual(expected, slidesTexts[1]);

            expected = new string[] { "Title 2", "Bullet 1", "Bullet 2", "Comment ça va ?" };
            CollectionAssert.AreEqual(expected, slidesTexts[2]);

            pptx.Close();
        }

        [TestMethod]
        public void GetSlides()
        {
            const string file = "../../files/GetTextsInAllSlides.pptx";

            Pptx pptx = new Pptx(file, false);
            int nbSlides = pptx.SlidesCount();
            Assert.AreEqual(3, nbSlides);

            IEnumerable<PptxSlide> slides = pptx.GetSlides();
            Assert.AreEqual(3, slides.Count());

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
                PptxSlide slide = pptx.GetSlide(i);
                string[] notes = slide.GetNotes();
                slidesNotes.Add(i, notes);
            }

            string[] expected = { "Bonjour", "{{comment1}}", "Hello", "1" };
            CollectionAssert.AreEqual(expected, slidesNotes[0]);

            expected = new string[] { "{{comment2}}", "2" };
            CollectionAssert.AreEqual(expected, slidesNotes[1]);

            expected = new string[] { };
            CollectionAssert.AreEqual(expected, slidesNotes[2]);

            // TODO Why "Comment çava ?" instead of "Comment ça va ?"
            expected = new string[] { "Bonjour {{comment3}} Hello", "Comment çava ?", string.Empty, string.Empty, "Hola!", string.Empty, "4" };
            CollectionAssert.AreEqual(expected, slidesNotes[3]);

            pptx.Close();
        }

        [TestMethod]
        public void FindSlides()
        {
            const string file = "../../files/GetNotesInAllSlides.pptx";

            Pptx pptx = new Pptx(file, false);
            int nbSlides = pptx.SlidesCount();
            Assert.AreEqual(4, nbSlides);

            IEnumerable<PptxSlide> slides = pptx.FindSlides("{{comment2}}");
            Assert.AreEqual(1, slides.Count());

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
                PptxSlide slide = pptx.GetSlide(i);
                PptxTable[] tables = slide.GetTables();
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
            {
                PptxSlide slide = pptx.GetSlide(0);
                slide.ReplaceTag("{{hello}}", "HELLO HOW ARE YOU?");
                slide.ReplaceTag("{{bonjour}}", "BONJOUR TOUT LE MONDE");
                slide.ReplaceTag("{{hola}}", "HOLA MAMA QUE TAL?");
            }

            // Second slide
            {
                PptxSlide slide = pptx.GetSlide(1);
                slide.ReplaceTag("{{hello}}", "H");
                slide.ReplaceTag("{{bonjour}}", "B");
                slide.ReplaceTag("{{hola}}", "H");
            }

            // Third slide
            {
                PptxSlide slide = pptx.GetSlide(2);
                slide.ReplaceTag("{{hello}}", string.Empty);
                slide.ReplaceTag("{{bonjour}}", string.Empty);
                slide.ReplaceTag("{{hola}}", null);
                slide.ReplaceTag(null, string.Empty);
                slide.ReplaceTag(null, null);
            }

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

            {
                PptxSlide slide = pptx.GetSlide(0);
                slide.ReplaceTag("{{hello}}", "HELLO!");
            }

            const string picture1_replace_png = "../../files/picture1_replace.png";
            const string picture1_replace_png_contentType = "image/png";
            const string picture1_replace_bmp = "../../files/picture1_replace.bmp";
            const string picture1_replace_bmp_contentType = "image/bmp";
            const string picture1_replace_jpeg = "../../files/picture1_replace.jpeg";
            const string picture1_replace_jpeg_contentType = "image/jpeg";
            const Stream picture1_replace_null = null;
            for (int i = 0; i < nbSlides; i++)
            {
                PptxSlide slide = pptx.GetSlide(i);

                slide.ReplacePicture("{{picture1png}}", picture1_replace_png, picture1_replace_png_contentType);
                slide.ReplacePicture("{{picture1bmp}}", picture1_replace_bmp, picture1_replace_bmp_contentType);
                slide.ReplacePicture("{{picture1jpeg}}", picture1_replace_jpeg, picture1_replace_jpeg_contentType);

                slide.ReplacePicture(null, picture1_replace_png, picture1_replace_png_contentType);
                slide.ReplacePicture("{{picture1null}}", picture1_replace_null, picture1_replace_png_contentType);
                slide.ReplacePicture("{{picture1null}}", picture1_replace_png, null);
                slide.ReplacePicture("{{picture1null}}", picture1_replace_null, null);
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

            PptxSlide slide = pptx.GetSlide(0);
            PptxTable table = slide.FindTables("{{table1}}").First();
            int[] columns = new int[] { 1, 3 };
            table.RemoveColumns(columns);

            pptx.Close();

            this.AssertPptxEquals(dstFileName, 1, "Column 0 Column2 Column 4 Cell 1.0 Cell 1.2 Cell 1.4 Cell 2.0 Cell 2.2 Cell 2.4 Cell 3.0 Cell 3.2 Cell 3.4 Cell 4.0 Cell 4.2 Cell 4.4 Cell 5.0 Cell 5.2 Cell 5.4 ");
        }

        [TestMethod]
        public void SetTableCellBackgroundPicture()
        {
            const string srcFileName = "../../files/SetTableCellBackgroundPicture.pptx";
            const string dstFileName = "../../files/SetTableCellBackgroundPicture_output.pptx";
            File.Delete(dstFileName);
            File.Copy(srcFileName, dstFileName);

            Pptx pptx = new Pptx(dstFileName, true);

            const string icon_png = "../../files/icon.png";
            const string icon_png_contentType = "image/png";
            FileStream iconFile = new FileStream(icon_png, FileMode.Open, FileAccess.Read);

            List<PptxTable.Cell[]> rows = new List<PptxTable.Cell[]>
                {
                    new[]
                        {
                            new PptxTable.Cell(
                                "{{cell0.0}}",
                                "Hello, world! 0.0",
                                new PptxTable.Cell.BackgroundPicture()
                                    {
                                        Picture = iconFile,
                                        ContentType = icon_png_contentType,
                                        Top = 14000,
                                        Right = 90000,
                                        Bottom = 12000,
                                        Left = 0
                                    }),
                            new PptxTable.Cell(
                                "{{cell3.0}}",
                                "Hello, world! 3.0",
                                new PptxTable.Cell.BackgroundPicture()
                                    {
                                        Picture = iconFile,
                                        ContentType = icon_png_contentType,
                                        Top = 14000,
                                        Right = 90000,
                                        Bottom = 12000,
                                        Left = 0
                                    })
                        },
                    new[]
                        {
                            new PptxTable.Cell(
                                "{{cell0.1}}",
                                "Hello, world! 0.1",
                                new PptxTable.Cell.BackgroundPicture()
                                    {
                                        Picture = iconFile,
                                        ContentType = icon_png_contentType,
                                        Top = 14000,
                                        Right = 90000,
                                        Bottom = 0,
                                        Left = 0
                                    })
                        },
                    new[]
                        {
                            new PptxTable.Cell(
                                "{{cell0.2}}",
                                "Hello, world! 0.2",
                                new PptxTable.Cell.BackgroundPicture()
                                    {
                                        Picture = iconFile,
                                        ContentType = icon_png_contentType,
                                        Top = 14000,
                                        Right = 0,
                                        Bottom = 0,
                                        Left = 0
                                    })
                        },
                    new[]
                        {
                            new PptxTable.Cell(
                                "{{cell0.3}}",
                                "Hello, world! 0.3",
                                new PptxTable.Cell.BackgroundPicture()
                                    {
                                        Picture = iconFile,
                                        ContentType = icon_png_contentType,
                                        Top = 0,
                                        Right = 0,
                                        Bottom = 0,
                                        Left = 0
                                    })
                        },
                    new[]
                        {
                            new PptxTable.Cell(
                                "{{cell0.4}}",
                                "Hello, world! 0.4",
                                new PptxTable.Cell.BackgroundPicture()
                                    {
                                        Picture = iconFile,
                                        ContentType = icon_png_contentType
                                    })
                        },
                    new[]
                        {
                            new PptxTable.Cell("{{cell0.5}}", "Hello, world! 0.5"),
                            new PptxTable.Cell("{{cell3.5}}", "Hello, world! 3.5")
                        }
                };

            {
                PptxSlide slide = pptx.GetSlide(0);
                PptxTable table = slide.FindTables("{{table1}}").First();
                table.SetRows(rows);
            }

            // Force a slide duplication using another table
            // This is to test that PptxSlide.Clone() works with background images
            {
                PptxSlide slide = pptx.GetSlide(0);
                PptxTable table = slide.FindTables("{{table2}}").First();
                table.SetRows(new List<PptxTable.Cell[]>());
            }

            pptx.Close();

            this.AssertPptxEquals(dstFileName, 1, "Col0 Col1 Col2 Col3 Col4 Hello, world! 0.0 Hello {{cell2.0}} Hello, world! 3.0 {{cell4.0}} Hello, world! 0.1 Hello {{cell2.1}} {{cell3.1}} {{cell4.1}} Hello, world! 0.2 Hello {{cell2.2}} {{cell3.2}} {{cell4.2}} Hello, world! 0.3 Hello {{cell2.3}} {{cell3.3}} {{cell4.3}} Hello, world! 0.4 Hello {{cell2.4}} {{cell3.4}} {{cell4.4}} Hello, world! 0.5 Hello {{cell2.5}} Hello, world! 3.5 {{cell4.5}}   ");
            // Sorry, you will have to manually check the background pictures
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
            {
                PptxSlide slide = pptx.GetSlide(0);
                slide.ReplaceTag("{{hello}}", "HELLO!");
            }

            // Change the pictures before to insert rows
            // PptxTable.SetRows() might change the number of slides inside the presentation
            {
                const string picture1_replace_png = "../../files/picture1_replace.png";
                const string picture1_replace_png_contentType = "image/png";
                PptxSlide slide = pptx.GetSlide(2);
                slide.ReplacePicture("{{picture1png}}", picture1_replace_png, picture1_replace_png_contentType);
            }

            {
                PptxSlide slide = pptx.GetSlide(0);
                PptxTable table = slide.FindTables("{{table1}}").First();
                List<PptxTable.Cell[]> rows = new List<PptxTable.Cell[]>();
                PptxTable.Cell[] row = new[]
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
                table.SetRows(rows);
            }

            List<PptxSlide> insertedSlides = new List<PptxSlide>();

            for (int i = 0; i < pptx.SlidesCount(); i++)
            {
                PptxSlide slide = pptx.GetSlide(i);
                PptxTable table = slide.FindTables("{{table2}}").FirstOrDefault();
                if (table != null)
                {
                    List<PptxTable.Cell[]> rows = new List<PptxTable.Cell[]>();
                    PptxTable.Cell[] row = new[]
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
                    insertedSlides.AddRange(table.SetRows(rows));
                }
            }

            foreach (PptxSlide insertedSlide in insertedSlides)
            {
                PptxTable table = insertedSlide.FindTables("{{table3}}").FirstOrDefault();
                if (table != null)
                {
                    List<PptxTable.Cell[]> rows = new List<PptxTable.Cell[]>();
                    PptxTable.Cell[] row = new[]
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
                    table.SetRows(rows);
                }
            }

            pptx.Close();

            this.AssertPptxEquals(dstFileName, 7, "Table1 Col2 Col3 Col4 Col5 Col6 HELLO Hello, world! 1  Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 HELLO Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 HELLO! Table1 Col2 Col3 Col4 Col5 Col6 HELLO Hello, world! 1  Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 HELLO Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 HELLO! Table1 Col2 Col3 Col4 Col5 Col6 HELLO Hello, world! 1  Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 Hello, world! 1 Hello, world! 2 Hello, world! 3 Hello, world! 4 Hello, world! 5 Hello, world! 6 HELLO! Table2 Col2 Col3 Col4 Col5 Col6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Table3 Col2 Col3 Col4 Col5 Col6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Table2 Col2 Col3 Col4 Col5 Col6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Table3 Col2 Col3 Col4 Col5 Col6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Table2 Col2 Col3 Col4 Col5 Col6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Bonjour 1 Bonjour 2 Bonjour 3 Bonjour 4 Bonjour 5 Bonjour 6 Table3 Col2 Col3 Col4 Col5 Col6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6 Hola! 1 Hola! 2 Hola! 3 Hola! 4 Hola! 5 Hola! 6  ");
        }

        [TestMethod]
        public void ReplaceTableMultipleTimes()
        {
            const string srcFileName = "../../files/ReplaceTableMultipleTimes.pptx";
            const string dstFileName = "../../files/ReplaceTableMultipleTimes_output.pptx";
            File.Delete(dstFileName);
            File.Copy(srcFileName, dstFileName);

            Pptx pptx = new Pptx(dstFileName, true);

            // Après la bataille (Victor Hugo)
            // http://fr.wikisource.org/wiki/Apr%C3%A8s_la_bataille_(Hugo)
            const string apresLaBataille =
                @"Mon père, ce héros au sourire si doux,
Suivi d’un seul housard qu’il aimait entre tous
Pour sa grande bravoure et pour sa haute taille,
Parcourait à cheval, le soir d’une bataille,
Le champ couvert de morts sur qui tombait la nuit.
Il lui sembla dans l’ombre entendre un faible bruit.
C’était un Espagnol de l’armée en déroute
Qui se traînait sanglant sur le bord de la route,
Râlant, brisé, livide, et mort plus qu’à moitié,
Et qui disait : « À boire ! à boire par pitié ! »
Mon père, ému, tendit à son housard fidèle
Une gourde de rhum qui pendait à sa selle,
Et dit : « Tiens, donne à boire à ce pauvre blessé. »
Tout à coup, au moment où le housard baissé
Se penchait vers lui, l’homme, une espèce de Maure,
Saisit un pistolet qu’il étreignait encore,
Et vise au front mon père en criant : « Caramba ! »
Le coup passa si près, que le chapeau tomba
Et que le cheval fit un écart en arrière.
« Donne-lui tout de même à boire », dit mon père.";

            // Le Dormeur du val (Arthur Rimbaud)
            // http://fr.wikisource.org/wiki/Le_Dormeur_du_val
            const string dormeurDuVal =
                @"C’est un trou de verdure où chante une rivière
Accrochant follement aux herbes des haillons
D’argent ; où le soleil, de la montagne fière,
Luit : c’est un petit val qui mousse de rayons.

Un soldat jeune, bouche ouverte, tête nue,
Et la nuque baignant dans le frais cresson bleu,
Dort ; il est étendu dans l’herbe, sous la nue,
Pâle dans son lit vert où la lumière pleut.

Les pieds dans les glaïeuls, il dort. Souriant comme
Sourirait un enfant malade, il fait un somme :
Nature, berce-le chaudement : il a froid.

Les parfums ne font pas frissonner sa narine ;
Il dort dans le soleil, la main sur sa poitrine
Tranquille. Il a deux trous rouges au côté droit.";

            List<List<string[]>> poems = new List<List<string[]>>();

            {
                string[] apresLaBatailleLines = apresLaBataille.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                List<string[]> lines = new List<string[]>();
                foreach (string line in apresLaBatailleLines)
                {
                    lines.Add(line.Split(new string[] { " " }, StringSplitOptions.None));
                }
                poems.Add(lines);
            }

            {
                string[] dormeurDuValLines = dormeurDuVal.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                List<string[]> lines = new List<string[]>();
                foreach (string line in dormeurDuValLines)
                {
                    lines.Add(line.Split(new string[] { " " }, StringSplitOptions.None));
                }
                poems.Add(lines);
            }

            {
                PptxSlide slideTemplate = pptx.GetSlide(0);
                PptxTable tableTemplate = slideTemplate.FindTables("{{table1}}").First();
                int rowsCountTemplate = tableTemplate.ColumnTitles().Length;

                PptxSlide prevSlide = slideTemplate;
                for (int i = 0; i < poems.Count; i++)
                {
                    PptxSlide slide = slideTemplate.Clone();
                    PptxSlide.InsertAfter(slide, prevSlide);
                    slide.ReplaceTag("{{title}}", i.ToString());

                    List<PptxTable.Cell[]> rows = new List<PptxTable.Cell[]>();

                    List<string[]> poem = poems[i];
                    foreach (string[] line in poem)
                    {
                        List<PptxTable.Cell> row = new List<PptxTable.Cell>();
                        for (int j = 0; j < rowsCountTemplate; j++)
                        {
                            PptxTable.Cell cell = new PptxTable.Cell("{{cell" + j + "}}", j < line.Length ? line[j] : string.Empty);
                            row.Add(cell);
                        }
                        rows.Add(row.ToArray());
                    }

                    PptxTable table = slide.FindTables("{{table1}}").First();
                    List<PptxSlide> insertedSlides = table.SetRows(rows);
                    PptxSlide lastInsertedSlide = insertedSlides.Last();
                    prevSlide = lastInsertedSlide;
                }

                slideTemplate.Remove();
            }

            pptx.Close();

            this.AssertPptxEquals(dstFileName, 6, "Col0 Col1 Col2 Col3 Col4 Col5 Col6 Col7 Col8 Col9 Col10 Col11 Col12 Col13 Mon père, ce héros au sourire si doux,       Suivi d’un seul housard qu’il aimait entre tous       Pour sa grande bravoure et pour sa haute taille,      Parcourait à cheval, le soir d’une bataille,        Le champ couvert de morts sur qui tombait la nuit.     Il lui sembla dans l’ombre entendre un faible bruit.      C’était un Espagnol de l’armée en déroute        0 Col0 Col1 Col2 Col3 Col4 Col5 Col6 Col7 Col8 Col9 Col10 Col11 Col12 Col13 Qui se traînait sanglant sur le bord de la route,     Râlant, brisé, livide, et mort plus qu’à moitié,       Et qui disait : « À boire ! à boire par pitié ! » Mon père, ému, tendit à son housard fidèle       Une gourde de rhum qui pendait à sa selle,      Et dit : « Tiens, donne à boire à ce pauvre blessé. »  Tout à coup, au moment où le housard baissé      0 Col0 Col1 Col2 Col3 Col4 Col5 Col6 Col7 Col8 Col9 Col10 Col11 Col12 Col13 Se penchait vers lui, l’homme, une espèce de Maure,      Saisit un pistolet qu’il étreignait encore,         Et vise au front mon père en criant : « Caramba ! »  Le coup passa si près, que le chapeau tomba      Et que le cheval fit un écart en arrière.      « Donne-lui tout de même à boire », dit mon père.    0 Col0 Col1 Col2 Col3 Col4 Col5 Col6 Col7 Col8 Col9 Col10 Col11 Col12 Col13 C’est un trou de verdure où chante une rivière      Accrochant follement aux herbes des haillons         D’argent ; où le soleil, de la montagne fière,      Luit : c’est un petit val qui mousse de rayons.                   Un soldat jeune, bouche ouverte, tête nue,        Et la nuque baignant dans le frais cresson bleu,      1 Col0 Col1 Col2 Col3 Col4 Col5 Col6 Col7 Col8 Col9 Col10 Col11 Col12 Col13 Dort ; il est étendu dans l’herbe, sous la nue,     Pâle dans son lit vert où la lumière pleut.                    Les pieds dans les glaïeuls, il dort. Souriant comme      Sourirait un enfant malade, il fait un somme :      Nature, berce-le chaudement : il a froid.                      1 Col0 Col1 Col2 Col3 Col4 Col5 Col6 Col7 Col8 Col9 Col10 Col11 Col12 Col13 Les parfums ne font pas frissonner sa narine ;      Il dort dans le soleil, la main sur sa poitrine     Tranquille. Il a deux trous rouges au côté droit.      1 ");
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

        [TestMethod]
        public void RemoveSlides()
        {
            const string srcFileName = "../../files/RemoveSlides.pptx";
            const string dstFileName = "../../files/RemoveSlides_output.pptx";
            File.Delete(dstFileName);
            File.Copy(srcFileName, dstFileName);

            Pptx pptx = new Pptx(dstFileName, true);
            Assert.AreEqual(5, pptx.SlidesCount());
            pptx.GetSlide(1).Remove();
            Assert.AreEqual(4, pptx.SlidesCount());
            pptx.Close();

            pptx = new Pptx(dstFileName, true);
            Assert.AreEqual(4, pptx.SlidesCount());
            pptx.GetSlide(0).Remove();
            pptx.GetSlide(2).Remove(); // 2 = 3 - the first slide removed
            Assert.AreEqual(2, pptx.SlidesCount());
            pptx.Close();

            File.Delete(dstFileName);
            File.Copy(srcFileName, dstFileName);
            pptx = new Pptx(dstFileName, true);
            int nbSlides = pptx.SlidesCount();
            Assert.AreEqual(5, nbSlides);
            for (int i = nbSlides - 1; i >= 0; i--)
            {
                if (i == 0 || i == 2)
                {
                    pptx.GetSlide(i).Remove();
                }
            }
            Assert.AreEqual(3, pptx.SlidesCount());
            pptx.Close();
        }

        [TestMethod]
        public void ReplaceTablesAndPictures()
        {
            const string srcFileName = "../../files/ReplaceTablesAndPictures.pptx";
            const string dstFileName = "../../files/ReplaceTablesAndPictures_output.pptx";
            File.Delete(dstFileName);
            File.Copy(srcFileName, dstFileName);

            Pptx pptx = new Pptx(dstFileName, true);

            List<PptxTable.Cell[]> rows = new List<PptxTable.Cell[]>
                {
                    new[] { new PptxTable.Cell("{{cell}}", "1") },
                    new[] { new PptxTable.Cell("{{cell}}", "2") },
                    new[] { new PptxTable.Cell("{{cell}}", "3") },
                    new[] { new PptxTable.Cell("{{cell}}", "4") },
                    new[] { new PptxTable.Cell("{{cell}}", "5") },
                    new[] { new PptxTable.Cell("{{cell}}", "6") }
                };

            PptxSlide slideTemplate = pptx.GetSlide(0);
            PptxSlide slide = slideTemplate.Clone();
            PptxSlide.InsertAfter(slide, slideTemplate);

            slide.ReplaceTag("{{hello}}", "Bonjour");

            const string picture1_replace_png = "../../files/picture1_replace.png";
            const string picture1_replace_png_contentType = "image/png";
            slide.ReplacePicture("{{picture1}}", picture1_replace_png, picture1_replace_png_contentType);

            PptxTable table1 = slide.FindTables("{{table1}}").First();
            List<PptxSlide> insertedSlides = table1.SetRows(rows);

            foreach (PptxSlide insertedSlide in insertedSlides)
            {
                PptxTable table2 = insertedSlide.FindTables("{{table2}}").First();
                table2.SetRows(rows);
            }

            slideTemplate.Remove();

            pptx.Close();

            this.AssertPptxEquals(dstFileName, 4, "Table1 1 2 3 4 Table2 1 2 3 4 Bonjour Table1 1 2 3 4 Table2 5 6 Bonjour Table1 5 6 Table2 1 2 3 4 Bonjour Table1 5 6 Table2 5 6 Bonjour ");
        }
    }
}

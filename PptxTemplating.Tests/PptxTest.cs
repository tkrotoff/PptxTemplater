using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PptxTemplating.Tests
{
    [TestClass]
    public class PptxTest
    {
        // See How to: Get All the Text in a Slide in a Presentation http://msdn.microsoft.com/en-us/library/office/cc850836
        // See How to: Get All the Text in All Slides in a Presentation http://msdn.microsoft.com/en-us/library/office/gg278331
        [TestMethod]
        public void TestGetAllTextInAllSlides()
        {
            string file = "../../files/test1.pptx";

            Pptx pptx = new Pptx(file, false);
            int nbSlides = pptx.CountSlides();
            Assert.AreEqual(3, nbSlides);

            var slidesText = new Dictionary<int, string[]>();
            for (int i = 0; i < nbSlides; i++)
            {
                string[] texts = pptx.GetAllTextInSlide(i);
                slidesText.Add(i, texts);
            }

            string[] expected = {"test1", "Hello, world!"};
            CollectionAssert.AreEqual(expected, slidesText[0]);
            expected = new string[] {"Titre 1", "Bullet 1", "Bullet 2"};
            CollectionAssert.AreEqual(expected, slidesText[1]);
            expected = new string[] {"Titre 2", "Bullet 1", "Bullet 2"};
            CollectionAssert.AreEqual(expected, slidesText[2]);
        }
    }
}

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
        // See How to: Get All the Text in All Slides in a Presentation http://msdn.microsoft.com/en-us/library/office/gg278331
        [TestMethod]
        public void TestGetAllTextInAllSlides()
        {
            string file = "../../files/test1.pptx";
            int numberOfSlides = Pptx.CountSlides(file);
            System.Console.WriteLine("Number of slides = {0}", numberOfSlides);
            string slideText;
            for (int i = 0; i < numberOfSlides; i++)
            {
                Pptx.GetSlideIdAndText(out slideText, file, i);
                System.Console.WriteLine("Slide #{0} contains: {1}", i + 1, slideText);
            }
        }
    }
}

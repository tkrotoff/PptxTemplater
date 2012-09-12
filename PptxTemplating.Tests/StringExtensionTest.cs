using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PptxTemplating.Tests
{
    [TestClass]
    public class StringExtensionTest
    {
        [TestMethod]
        public void TestSubstrings()
        {
            string str = "";
            int[] splits = {};
            string[] substrings = {};

            // Regular case, last split too small compared to the given string
            str = "Bonjour tout le monde";
            splits = new int[] { 5, 8, 3 };
            substrings = str.Substrings(splits);
            Assert.AreEqual(3, substrings.Length);
            Assert.AreEqual("Bonjo", substrings[0]);
            Assert.AreEqual("ur tout ", substrings[1]);
            Assert.AreEqual("le ", substrings[2]);

            // Last split too big compared to the given string
            str = "Bonjour tout le monde";
            splits = new int[] { 5, 8, 100 };
            substrings = str.Substrings(splits);
            Assert.AreEqual(3, substrings.Length);
            Assert.AreEqual("Bonjo", substrings[0]);
            Assert.AreEqual("ur tout ", substrings[1]);
            Assert.AreEqual("le monde", substrings[2]);

            // Middle split too big compared to the given string
            str = "Bonjour tout le monde";
            splits = new int[] { 5, 100, 3 };
            substrings = str.Substrings(splits);
            Assert.AreEqual(3, substrings.Length);
            Assert.AreEqual("Bonjo", substrings[0]);
            Assert.AreEqual("ur tout le monde", substrings[1]);
            Assert.AreEqual("", substrings[2]);

            // Too many splits compared to the given string
            str = "Bonjour tout le monde";
            splits = new int[] { 5, 8, 3, 5, 5, 5 };
            substrings = str.Substrings(splits);
            Assert.AreEqual(6, substrings.Length);
            Assert.AreEqual("Bonjo", substrings[0]);
            Assert.AreEqual("ur tout ", substrings[1]);
            Assert.AreEqual("le ", substrings[2]);
            Assert.AreEqual("monde", substrings[3]);
            Assert.AreEqual("", substrings[4]);
            Assert.AreEqual("", substrings[5]);

            // Split too big compared to the given string
            str = "Bonjour tout le monde";
            splits = new int[] { 100 };
            substrings = str.Substrings(splits);
            Assert.AreEqual(1, substrings.Length);
            Assert.AreEqual("Bonjour tout le monde", substrings[0]);

            // Empty split
            str = "Bonjour tout le monde";
            splits = new int[] { 0 };
            substrings = str.Substrings(splits);
            Assert.AreEqual(1, substrings.Length);
            Assert.AreEqual("", substrings[0]);

            // No split
            str = "Bonjour tout le monde";
            splits = new int[] { };
            substrings = str.Substrings(splits);
            Assert.AreEqual(0, substrings.Length);
        }
    }
}

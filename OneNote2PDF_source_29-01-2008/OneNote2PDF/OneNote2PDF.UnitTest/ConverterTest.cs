using OneNote2PDF.Library;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace OneNote2PDF.UnitTest
{
    
    
    /// <summary>
    ///This is a test class for ConverterTest and is intended
    ///to contain all ConverterTest Unit Tests
    ///</summary>
    [TestClass()]
    public class ConverterTest
    {
        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        // 
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        /// <summary>
        ///A test for NumberToRoman
        ///</summary>
        [TestMethod()]
        public void NumberToRomanTest()
        {
            int number = 0;
            bool lowerCase = false;
            string expected = "N";
            string actual;
            actual = Converter.NumberToRoman(number, lowerCase);
            Assert.AreEqual(expected, actual);

            number = 10;
            lowerCase = false;
            expected = "X";
            actual = Converter.NumberToRoman(number, lowerCase);
            Assert.AreEqual(expected, actual);

            number = 8;
            lowerCase = false;
            expected = "VIII";
            actual = Converter.NumberToRoman(number, lowerCase);
            Assert.AreEqual(expected, actual);

            number = 2000;
            lowerCase = false;
            expected = "MM";
            actual = Converter.NumberToRoman(number, lowerCase);
            Assert.AreEqual(expected, actual);

            number = 19;
            lowerCase = false;
            expected = "XIX";
            actual = Converter.NumberToRoman(number, lowerCase);
            Assert.AreEqual(expected, actual);
        }
    }
}

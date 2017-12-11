using OneNote2PDF.Library;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace OneNote2PDF.UnitTest
{
    
    
    /// <summary>
    ///This is a test class for TOCHandlerTest and is intended
    ///to contain all TOCHandlerTest Unit Tests
    ///</summary>
    [TestClass()]
    public class TOCHandlerTest
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
        [ClassInitialize()]
        public static void MyClassInitialize(TestContext testContext)
        {
            Config.Current.TraceLevel = System.Diagnostics.TraceLevel.Off;
        }
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
        /// Test the login of adding toc at different levels
        /// </summary>
        [TestMethod()]
        public void LogicTest()
        {
            TOCHandler_Accessor target = new TOCHandler_Accessor();
            target.Init(null, null);
            target.BeginTocEntry(); // level 1
            Assert.AreEqual(target.currentLevel, 1);
            Assert.AreEqual(target.TOCNumbering[1], 1);
            target.AddTocEntry("hello", 10);
            target.AddTocEntry("hello", 10);
            target.AddTocEntry("hello", 10); // toc numbering should equal to 3 + 1
            Assert.AreEqual(target.TOCNumbering[1], 4);

            target.BeginTocEntry(); // increase level to 2
            target.AddTocEntry("hello", 10);
            target.AddTocEntry("hello", 10); // toc numbering should equal to 2 + 1
            Assert.AreEqual(target.TOCNumbering[1], 4);
            Assert.AreEqual(target.TOCNumbering[2], 3);

            // add empty level
            target.BeginTocEntry(); // increase level to 3
            target.BeginTocEntry(); // increase level to 4
            target.AddTocEntry("hello", 10);
            Assert.AreEqual(target.TOCNumbering[3], 1);
            Assert.AreEqual(target.TOCNumbering[4], 2);
            target.EndTocEntry();
            target.EndTocEntry();

            target.EndTocEntry();
            target.EndTocEntry();
            Assert.AreEqual(target.currentLevel, 0);
        }
        /// <summary>
        ///A test for Init
        ///</summary>
        [TestMethod()]
        public void InitTest()
        {
            TOCHandler target = new TOCHandler(); // TODO: Initialize to an appropriate value
            Document document = null; // TODO: Initialize to an appropriate value
            PdfWriter writer = null; // TODO: Initialize to an appropriate value
            target.Init(document, writer);
        }

        /// <summary>
        ///A test for EndTocEntry
        ///</summary>
        [TestMethod(), ExpectedException(typeof(System.ArgumentOutOfRangeException))]
        public void EndTocLevel0Test()
        {
            TOCHandler_Accessor target = new TOCHandler_Accessor();
            target.Init(null, null);
            target.BeginTocEntry();
            target.EndTocEntry();
            target.EndTocEntry();
            target.EndTocEntry();
        }

        /// <summary>
        ///A test for EndTocEntry
        ///</summary>
        [TestMethod(), ExpectedException(typeof(System.ArgumentOutOfRangeException))]
        public void EndTocLevel1Test()
        {
            TOCHandler_Accessor target = new TOCHandler_Accessor();
            target.EndTocEntry();
        }

        /// <summary>
        ///A test for EndTocEntry
        ///</summary>
        [TestMethod(), ExpectedException(typeof(System.ArgumentOutOfRangeException))]
        public void EndTocLevel2Test()
        {
            TOCHandler_Accessor target = new TOCHandler_Accessor();
            target.Init(null, null);
            target.EndTocEntry();
        }

        /// <summary>
        ///A test for BeginTocEntry
        ///</summary>
        [TestMethod(), ExpectedException(typeof(System.InvalidOperationException))]
        public void BeginTocEntry1Test()
        {
            TOCHandler_Accessor target = new TOCHandler_Accessor();
            target.BeginTocEntry();
        }
        /// <summary>
        ///A test for BeginTocEntry
        ///</summary>
        [TestMethod()]
        public void BeginTocEntry2Test()
        {
            TOCHandler_Accessor target = new TOCHandler_Accessor();
            int level = 1;
            target.Init(null, null);
            target.BeginTocEntry();
            Assert.AreEqual(target.currentLevel, level);
            Assert.AreEqual(target.TOCNumbering[level], 1);
        }
        /// <summary>
        ///A test for AddTocEntry
        ///</summary>
        [TestMethod()]
        public void AddTocEntryTest()
        {
            TOCHandler_Accessor target = new TOCHandler_Accessor();
            int level = 1;
            target.Init(null, null);
            target.BeginTocEntry();
            target.AddTocEntry("title", 10);
            Assert.AreEqual(target.TOCNumbering[level], 2);
        }

        /// <summary>
        ///A test for AddTocEntry
        ///</summary>
        [TestMethod(), ExpectedException(typeof(System.InvalidOperationException))]
        public void AddTocEntry1Test()
        {
            TOCHandler_Accessor target = new TOCHandler_Accessor();
            target.AddTocEntry("title", 10);
        }
        [TestMethod(), ExpectedException(typeof(System.InvalidOperationException))]
        public void AddTocEntry2Test()
        {
            TOCHandler_Accessor target = new TOCHandler_Accessor();
            target.Init(null, null);
            target.AddTocEntry("title", 10);
        }
        [TestMethod()]
        public void AddTocEntry3Test()
        {
            TOCHandler_Accessor target = new TOCHandler_Accessor();
            int level = 1;
            target.Init(null, null);
            target.BeginTocEntry();
            target.AddTocEntry("title", 10);
            Assert.AreEqual(level, target.currentLevel);
            Assert.AreEqual(target.TOCNumbering[level], 2);
        }
    }
}

using OneNote2PDF.Library;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OneNote2PDF.Library.Data;
using iTextSharp.text.pdf;
using Microsoft.Office.Interop.OneNote;
using iTextSharp.text;
using System.Collections.Generic;

namespace OneNote2PDF.UnitTest
{
    
    
    /// <summary>
    ///This is a test class for Export2PDFTest and is intended
    ///to contain all Export2PDFTest Unit Tests
    ///</summary>
    [TestClass()]
    public class Export2PDFTest
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
        ///A test for PDFMergeAll
        ///</summary>
        public void PDFMergeAllTestHelper<T>()
        {
            Export2PDF_Accessor target = new Export2PDF_Accessor(); // TODO: Initialize to an appropriate value
            string basedName = string.Empty; // TODO: Initialize to an appropriate value
            ONSection sec = null; // TODO: Initialize to an appropriate value
            target.PDFMergeAll(basedName, sec);
            Assert.Inconclusive("A method that does not return a value cannot be verified.");
        }

        [TestMethod(), ExpectedException(typeof(System.ArgumentNullException))]
        [DeploymentItem("OneNote2PDF.exe")]
        public void PDFMergeAllTest()
        {
            PDFMergeAllTestHelper<GenericParameterHelper>();
        }

        /// <summary>
        ///A test for PDFExportSection
        ///</summary>
        [TestMethod(), ExpectedException(typeof(System.ArgumentNullException))]
        [DeploymentItem("OneNote2PDF.exe")]
        public void PDFExportSectionTest()
        {
            Export2PDF_Accessor target = new Export2PDF_Accessor(); // TODO: Initialize to an appropriate value
            string basedName = string.Empty; // TODO: Initialize to an appropriate value
            ONSection section = null; // TODO: Initialize to an appropriate value
            target.PDFExportSection(basedName, section);
        }

        /// <summary>
        ///A test for PDFExportPage
        ///</summary>
        [TestMethod(), ExpectedException(typeof(System.ArgumentNullException))]
        [DeploymentItem("OneNote2PDF.exe")]
        public void PDFExportPageTest()
        {
            Export2PDF_Accessor target = new Export2PDF_Accessor();
            string basedName = string.Empty;
            ONPage page = null;
            target.PDFExportPage(basedName, page);
        }

        /// <summary>
        ///A test for PDFCombineSection
        ///</summary>
        [TestMethod(), ExpectedException(typeof(System.ArgumentNullException))]
        [DeploymentItem("OneNote2PDF.exe")]
        public void PDFCombineSectionTest()
        {
            Export2PDF_Accessor target = new Export2PDF_Accessor(); 
            PdfOutline parent = null; 
            ONSection section = null;
            target.PDFCombineSection(parent, section);
        }

        /// <summary>
        ///A test for PDFCombinePage
        ///</summary>
        [TestMethod(), ExpectedException(typeof(System.ArgumentNullException))]
        [DeploymentItem("OneNote2PDF.exe")]
        public void PDFCombinePageTest()
        {
            Export2PDF_Accessor target = new Export2PDF_Accessor();
            PdfOutline parent = null; 
            ONPage page = null; 
            target.PDFCombinePage(parent, page);
        }

        /// <summary>
        ///A test for PDFCombineAll
        ///</summary>
        [TestMethod(), ExpectedException(typeof(System.ArgumentNullException))]
        [DeploymentItem("OneNote2PDF.exe")]
        public void PDFCombineAllTest()
        {
            Export2PDF_Accessor target = new Export2PDF_Accessor(); 
            string basedName = string.Empty; 
            ONNotebook notebook = null; 
            target.PDFCombineAll(basedName, notebook);
        }

        /// <summary>
        ///A test for InitDocument
        ///</summary>
        [TestMethod()]
        [DeploymentItem("OneNote2PDF.exe")]
        public void InitDocumentTest()
        {
            Export2PDF_Accessor target = new Export2PDF_Accessor(); // TODO: Initialize to an appropriate value
            string pathName = string.Empty; // TODO: Initialize to an appropriate value
            bool expected = false; // TODO: Initialize to an appropriate value
            bool actual;
            actual = target.InitDocument(pathName);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for Export
        ///</summary>
        public void ExportTestHelper<T>() where T: ONBasedType
        {
            Export2PDF target = new Export2PDF(); // TODO: Initialize to an appropriate value
            string pathName = string.Empty; // TODO: Initialize to an appropriate value
            T part = default(T); // TODO: Initialize to an appropriate value
            target.Export<T>(pathName, part);
        }

        [TestMethod()]
        public void ExportTest()
        {
            ExportTestHelper<ONBasedType>();
        }

        /// <summary>
        ///A test for CloseDocument
        ///</summary>
        [TestMethod()]
        [DeploymentItem("OneNote2PDF.exe")]
        public void CloseDocumentTest()
        {
            Export2PDF_Accessor target = new Export2PDF_Accessor(); // TODO: Initialize to an appropriate value
            target.CloseDocument();
            Assert.IsNull(target.pdfDocument);
        }


        /// <summary>
        ///A test for Export2PDF Constructor
        ///</summary>
        [TestMethod()]
        public void Export2PDFConstructorTest()
        {
            Export2PDF target = new Export2PDF();
            Assert.IsNull(target.OneNoteApplication);
            Assert.IsTrue(string.IsNullOrEmpty(target.BasedPath));
        }
    }
}

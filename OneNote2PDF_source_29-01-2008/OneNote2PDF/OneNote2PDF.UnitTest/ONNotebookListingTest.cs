using OneNote2PDF.Library;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OneNote2PDF.Library.Data;
using System.Xml.Linq;
using System.Collections.Generic;

namespace OneNote2PDF.UnitTest
{
    
    
    /// <summary>
    ///This is a test class for ONNotebookListingTest and is intended
    ///to contain all ONNotebookListingTest Unit Tests
    ///</summary>
    [TestClass()]
    public class ONNotebookListingTest
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
        ///A test for SelectSections
        ///</summary>
        [TestMethod()]
        [DeploymentItem("OneNote2PDF.exe")]
        public void SelectSectionsTest()
        {
            string notebooksXml = @"<?xml version=""1.0"" encoding=""utf-8"" ?>
<one:Notebooks xmlns:one=""http://schemas.microsoft.com/office/onenote/2007/onenote"">
  <one:Notebook name=""Code Notebook"" ID=""{B4B9B6D8-C298-4CF1-A191-C30C5BF0B69D}{1}{B0}"" lastModifiedTime=""2007-12-07T10:15:56.000Z"">
    <one:Section name=""TODO"" ID=""{928433B9-7EE1-0C1D-38FB-21B387CEFA94}{1}{B0}"" lastModifiedTime=""2007-08-31T14:53:10.000Z"">
        <one:Page name=""Test"" ID=""{928433B9-7EE1-0C1D-38FB-21B387CEFA94}{1}{B0}"" lastModifiedTime=""2007-08-31T14:53:10.000Z"" />
        </one:Section>
    </one:Notebook>
</one:Notebooks>";
            ONNotebookListing_Accessor target = new ONNotebookListing_Accessor(notebooksXml);
            XDocument xmlDoc = XDocument.Parse(notebooksXml);
            XNamespace ns = "http://schemas.microsoft.com/office/onenote/2007/onenote";
            IEnumerator<XElement> xmlItr = xmlDoc.Descendants(ns + "Notebook").GetEnumerator();
            xmlItr.MoveNext();
            XElement xml = xmlItr.Current;
            List<ONSection> actual;
            actual = target.SelectSections(xml);
            string expected = "TODO";
            Assert.IsNotNull(actual);
            Assert.AreEqual(expected, actual[0].Name);
        }

        /// <summary>
        ///A test for SelectPages
        ///</summary>
        [TestMethod()]
        [DeploymentItem("OneNote2PDF.exe")]
        public void SelectPagesTest()
        {
            string notebooksXml = @"<?xml version=""1.0"" encoding=""utf-8"" ?>
<one:Notebooks xmlns:one=""http://schemas.microsoft.com/office/onenote/2007/onenote"">
  <one:Notebook name=""Code Notebook"" ID=""{B4B9B6D8-C298-4CF1-A191-C30C5BF0B69D}{1}{B0}"" lastModifiedTime=""2007-12-07T10:15:56.000Z"">
    <one:Section name=""TODO"" ID=""{928433B9-7EE1-0C1D-38FB-21B387CEFA94}{1}{B0}"" lastModifiedTime=""2007-08-31T14:53:10.000Z"">
        <one:Page name=""Test"" ID=""{928433B9-7EE1-0C1D-38FB-21B387CEFA94}{1}{B0}"" lastModifiedTime=""2007-08-31T14:53:10.000Z"" />
        </one:Section>
    </one:Notebook>
</one:Notebooks>";
            ONNotebookListing_Accessor target = new ONNotebookListing_Accessor(notebooksXml);
            XDocument xmlDoc = XDocument.Parse(notebooksXml);
            XNamespace ns = "http://schemas.microsoft.com/office/onenote/2007/onenote";
            IEnumerator<XElement> xmlItr = xmlDoc.Descendants(ns + "Section").GetEnumerator();
            xmlItr.MoveNext();
            XElement xml = xmlItr.Current;
            List<ONPage> actual;
            actual = target.SelectPages(xml);
            string expected = "Test";
            Assert.IsNotNull(actual);
            Assert.AreEqual(expected, actual[0].Name);
        }

        /// <summary>
        ///A test for ListAllNotebook
        ///</summary>
        [TestMethod()]
        public void ListAllNotebookTest()
        {
            string notebooksXml = @"<?xml version=""1.0"" encoding=""utf-8"" ?>
<one:Notebooks xmlns:one=""http://schemas.microsoft.com/office/onenote/2007/onenote"">
  <one:Notebook name=""Code Notebook"" ID=""{B4B9B6D8-C298-4CF1-A191-C30C5BF0B69D}{1}{B0}"" lastModifiedTime=""2007-12-07T10:15:56.000Z"">
    </one:Notebook>
  <one:Notebook name=""Code Notebook1"" ID=""{B4B9B6D8-C298-4CF1-A191-C30C5BF0B69D}{1}{B0}"" lastModifiedTime=""2007-12-07T10:15:56.000Z"">
    </one:Notebook>
  <one:Notebook name=""Code Notebook2"" ID=""{B4B9B6D8-C298-4CF1-A191-C30C5BF0B69D}{1}{B0}"" lastModifiedTime=""2007-12-07T10:15:56.000Z"">
    </one:Notebook>
</one:Notebooks>";
            ONNotebookListing target = new ONNotebookListing(notebooksXml);
            string[] expected = new string[] {"Code Notebook", "Code Notebook1", "Code Notebook2"};
            string[] actual;
            actual = target.ListAllNotebook();
            CollectionAssert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for GetNotebook
        ///</summary>
        [TestMethod()]
        public void GetNotebookTest()
        {
            string notebooksXml = @"<?xml version=""1.0"" encoding=""utf-8"" ?>
<one:Notebooks xmlns:one=""http://schemas.microsoft.com/office/onenote/2007/onenote"">
  <one:Notebook name=""Code Notebook"" ID=""{B4B9B6D8-C298-4CF1-A191-C30C5BF0B69D}{1}{B0}"" lastModifiedTime=""2007-12-07T10:15:56.000Z"">
    <one:Section name=""TODO"" ID=""{928433B9-7EE1-0C1D-38FB-21B387CEFA94}{1}{B0}"" lastModifiedTime=""2007-08-31T14:53:10.000Z"">
        <one:Page name=""Test"" ID=""{928433B9-7EE1-0C1D-38FB-21B387CEFA94}{1}{B0}"" lastModifiedTime=""2007-08-31T14:53:10.000Z"" />
        </one:Section>
    </one:Notebook>
</one:Notebooks>";
            ONNotebookListing target = new ONNotebookListing(notebooksXml); // TODO: Initialize to an appropriate value
            string notebookName = string.Empty; // TODO: Initialize to an appropriate value
            ONNotebook expected = null; // TODO: Initialize to an appropriate value
            ONNotebook actual;
            actual = target.GetNotebook(notebookName);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void GetNotebook1Test()
        {
            string notebooksXml = @"<?xml version=""1.0"" encoding=""utf-8"" ?>
<one:Notebooks xmlns:one=""http://schemas.microsoft.com/office/onenote/2007/onenote"">
  <one:Notebook name=""Code Notebook"" ID=""{B4B9B6D8-C298-4CF1-A191-C30C5BF0B69D}{1}{B0}"" lastModifiedTime=""2007-12-07T10:15:56.000Z"">
    <one:Section name=""TODO"" ID=""{928433B9-7EE1-0C1D-38FB-21B387CEFA94}{1}{B0}"" lastModifiedTime=""2007-08-31T14:53:10.000Z"">
        <one:Page name=""Test"" ID=""{928433B9-7EE1-0C1D-38FB-21B387CEFA94}{1}{B0}"" lastModifiedTime=""2007-08-31T14:53:10.000Z"" />
        </one:Section>
    </one:Notebook>
</one:Notebooks>";

            ONNotebookListing target = new ONNotebookListing(notebooksXml);
            string expect = "Code Notebook";
            ONNotebook actual = target.GetNotebook("Code Notebook");
            Assert.AreEqual(expect, actual.Name);
        }

        [TestMethod()]
        public void GetNotebook2Test()
        {
            string notebooksXml = @"<?xml version=""1.0"" encoding=""utf-8"" ?>
<one:Notebooks xmlns:one=""http://schemas.microsoft.com/office/onenote/2007/onenote"">
  <one:Notebook name=""Code Notebook"" ID=""{B4B9B6D8-C298-4CF1-A191-C30C5BF0B69D}{1}{B0}"" lastModifiedTime=""2007-12-07T10:15:56.000Z"">
    <one:Section name=""TODO"" ID=""{928433B9-7EE1-0C1D-38FB-21B387CEFA94}{1}{B0}"" lastModifiedTime=""2007-08-31T14:53:10.000Z"">
        <one:Page name=""Test"" ID=""{928433B9-7EE1-0C1D-38FB-21B387CEFA94}{1}{B0}"" lastModifiedTime=""2007-08-31T14:53:10.000Z"" />
        </one:Section>
    </one:Notebook>
</one:Notebooks>";

            ONNotebookListing target = new ONNotebookListing(notebooksXml);
            string expect = "Code Notebook";
            ONNotebook actual = target.GetNotebook("code notebook");
            Assert.AreEqual(expect, actual.Name);
        }

        [TestMethod()]
        public void GetNotebook3Test()
        {
            string notebooksXml = @"<?xml version=""1.0"" encoding=""utf-8"" ?>
<one:Notebooks xmlns:one=""http://schemas.microsoft.com/office/onenote/2007/onenote"">
  <one:Notebook name=""Code Notebook"" ID=""{B4B9B6D8-C298-4CF1-A191-C30C5BF0B69D}{1}{B0}"" lastModifiedTime=""2007-12-07T10:15:56.000Z"">
    <one:Section name=""TODO"" ID=""{928433B9-7EE1-0C1D-38FB-21B387CEFA94}{1}{B0}"" lastModifiedTime=""2007-08-31T14:53:10.000Z"">
        <one:Page name=""Test"" ID=""{928433B9-7EE1-0C1D-38FB-21B387CEFA94}{1}{B0}"" lastModifiedTime=""2007-08-31T14:53:10.000Z"" />
        </one:Section>
    </one:Notebook>
</one:Notebooks>";

            ONNotebookListing target = new ONNotebookListing(notebooksXml);
            ONNotebook expect = null;
            ONNotebook actual = target.GetNotebook("asdaosds");
            Assert.AreEqual(expect, actual);
        }

        /// <summary>
        ///A test for ONNotebookListing Constructor
        ///</summary>
        [TestMethod(), ExpectedException(typeof(System.Xml.XmlException))]
        public void ONNotebookListingConstructorTest()
        {
            string notebooksXml = string.Empty; // TODO: Initialize to an appropriate value
            ONNotebookListing target = new ONNotebookListing(notebooksXml);
        }

        /// <summary>
        ///A test for ONNotebookListing Constructor
        ///</summary>
        [TestMethod(), ExpectedException(typeof(System.Xml.XmlException))]
        public void ONNotebookListingConstructor1Test()
        {
            string notebooksXml = "something invalid for XML";
            ONNotebookListing target = new ONNotebookListing(notebooksXml);
        }

        /// <summary>
        ///A test for ONNotebookListing Constructor
        ///</summary>
        [TestMethod()]
        public void ONNotebookListingConstructor2Test()
        {
            string notebooksXml = @"<?xml version=""1.0"" encoding=""utf-8"" ?>
<one:Notebooks xmlns:one=""http://schemas.microsoft.com/office/onenote/2007/onenote"">
  <one:Notebook name=""Code Notebook"" ID=""{B4B9B6D8-C298-4CF1-A191-C30C5BF0B69D}{1}{B0}"" lastModifiedTime=""2007-12-07T10:15:56.000Z"">
    <one:Section name=""TODO"" ID=""{928433B9-7EE1-0C1D-38FB-21B387CEFA94}{1}{B0}"" lastModifiedTime=""2007-08-31T14:53:10.000Z"">
        <one:Page name=""Test"" ID=""{928433B9-7EE1-0C1D-38FB-21B387CEFA94}{1}{B0}"" lastModifiedTime=""2007-08-31T14:53:10.000Z"" />
        </one:Section>
    </one:Notebook>
</one:Notebooks>";

            ONNotebookListing target = new ONNotebookListing(notebooksXml);
            string expect = "Code Notebook";
            ONNotebook actual = target.GetNotebook("Code Notebook");
            Assert.AreEqual(expect, actual.Name);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OneNote2PDF.Library
{
    class ONNotebookListing
    {
        #region Private const
        private const string namespaceOneNote2007 = "http://schemas.microsoft.com/office/onenote/2007/onenote";
        // Define XML namespace
        XNamespace oneNS = namespaceOneNote2007;
        private string OneNoteNoteBookXML;
        private List<Data.ONNotebook> listNotebooks;
        #endregion

        #region Constructor and destructor

        public ONNotebookListing(string notebooksXml)
        {
            OneNoteNoteBookXML = notebooksXml;

            PopulateData();

        }

        private void PopulateData()
        {
            try
            {
                XDocument outputXML = XDocument.Parse(OneNoteNoteBookXML);
                listNotebooks = (from notebook in outputXML.Descendants(oneNS + "Notebook")
                                 select new Data.ONNotebook
                                 {
                                     Name = notebook.Attribute("name").Value,
                                     ID = notebook.Attribute("ID").Value,
                                     LastModifiedTime = Convert.ToDateTime(notebook.Attribute("lastModifiedTime").Value),
                                     Sections = SelectSections(notebook),
                                 }).ToList();

                foreach (Data.ONNotebook nb in listNotebooks)
                {
                    foreach (Data.ONSection s in nb.Sections)
                    {
                        s.Parent = nb;
                    }
                }
            }
            catch (System.Xml.XmlException ex)
            {
                Log.Error(ex.Message);
                throw;
            }
            catch (InvalidOperationException ex)
            {
                // source is empty
                Log.Error(ex.Message);
                throw;
            }
        }
        ~ONNotebookListing()
        {

        }

        #endregion

        #region Public methods

        public string[] ListAllNotebook()
        {
            if (listNotebooks == null)
                throw new InvalidOperationException("Object is used before initialized");

            string[] notebookNames = (from notebook in listNotebooks
                                      select notebook.Name).ToArray();

            return notebookNames;
        }

        public Data.ONNotebook GetNotebook(string notebookName)
        {
            try
            {
                Data.ONNotebook notebook = (from nb in listNotebooks
                                            where nb.Name.ToLower() == notebookName.ToLower()
                                            select nb).First();
                return notebook;
            }
            catch (InvalidOperationException)
            {
                return null;
            }
            catch (ArgumentNullException)
            {
                return null;
            }

        }
        #endregion

        #region LINQ Data Manipulation
        private List<Data.ONSection> SelectSections(XElement xml)
        {
            List<Data.ONSection> sections =
                (from section in xml.Elements()
                 where
                 (section.Name == oneNS + "SectionGroup") || (section.Name == oneNS + "Section")
                 // orderby section.Value
                 select new Data.ONSection
                 {
                     Name = section.Attribute("name").Value,
                     ID = section.Attribute("ID").Value,
                     LastModifiedTime = Convert.ToDateTime(section.Attribute("lastModifiedTime").Value),
                     SubSections = SelectSections(section),
                     Pages = SelectPages(section),
                 }).ToList();
            foreach (Data.ONSection s in sections)
            {
                foreach (Data.ONSection ss in s.SubSections)
                {
                    ss.Parent = s;
                }
                foreach (Data.ONPage p in s.Pages)
                {
                    p.Parent = s;
                }
            }
            return sections;
        }
        // Get page information out of XML
        private List<Data.ONPage> SelectPages(XElement xml)
        {
            List<Data.ONPage> pages =
                (from page in xml.Elements(oneNS + "Page")
                 // orderby section.Value
                 select new Data.ONPage
                 {
                     Name = page.Attribute("name").Value,
                     ID = page.Attribute("ID").Value,
                     LastModifiedTime = Convert.ToDateTime(page.Attribute("lastModifiedTime").Value),
                 }).ToList();
            return pages;
        }

        private Data.ONNotebook SelectNotebook(string notebookName, string xml)
        {
            try
            {
                XDocument outputXML = XDocument.Parse(xml);

                Data.ONNotebook notebooks = (from notebook in outputXML.Descendants(oneNS + "Notebook")
                                             where notebook.Attribute("name").Value == notebookName
                                             select new Data.ONNotebook
                                             {
                                                 Name = notebook.Attribute("name").Value,
                                                 ID = notebook.Attribute("ID").Value,
                                                 LastModifiedTime = Convert.ToDateTime(notebook.Attribute("lastModifiedTime").Value),
                                                 Sections = SelectSections(notebook),
                                             }).First();
                return notebooks;
            }
            catch (InvalidOperationException)
            {
                // source is empty
                return null;
            }
            catch (ArgumentNullException)
            {
                // source is null
                return null;
            }
            catch (System.Xml.XmlException ex)
            {
                // problem with input XML 
                Log.Error(ex.Message);
                return null;
            }
        }

        #endregion
    }
}

using System;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using OneNote2PDF.Library.Data;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace OneNote2PDF.Library
{
    class Export2PDF
    {
        #region Private area

        #region Private variables
        private Document pdfDocument { get; set; }
        private PdfWriter pdfWriter { get; set; }
        private TOCHandler TOCHandler { get; set; }
        #endregion

        #endregion

        #region Public area

        #region Public properties
        public OneNote.Application OneNoteApplication { get; set; }
        public string BasedPath { get; set; }
        #endregion

        #region Public methods
        public Export2PDF()
        {
            TOCHandler = new TOCHandler();
        }

        public void CreateCacheFolder(Data.ONNotebook notebook)
        {
            if (string.IsNullOrEmpty(BasedPath))
            {
                Log.Error("BasedPath cannot be null");
                return;
            }
            if (notebook==null)
            {
                Log.Error("Notebook cannot be null");
                throw new ArgumentNullException("notebook");
            }
            string baseDir = Path.Combine(BasedPath, notebook.Name);
            Directory.CreateDirectory(baseDir);

            if (Config.Current.RefreshCache)
            {
                Log.Information("Refreshing the cache folder");
                // delete entire cache folder
                Directory.Delete(baseDir, true);
            }
            SetIncludeOnlySection(notebook);
            SetExcludeSection(notebook);

            foreach (Data.ONSection section in notebook.Sections)
            {
                if (!section.Exclude)
                {
                    PDFExportSection(baseDir, section);
                }
            }
        }

        public void Export<T>(string pathName, T part) where T : ONBasedType
        {
            if (string.IsNullOrEmpty(pathName))
            {
                Log.Error("PathName cannot be null");
                return;
            }
            if (part == null)
            {
                Log.Error("Part cannot be null");
                return;
            }
            if (part is ONNotebook)
            {
                PDFCombineAll(pathName, part as ONNotebook);
            }
            else
                if (part is ONSection)
                {
                    PDFMergeAll(pathName, part as ONSection);
                }
        }

        #endregion

        #endregion

        #region Private methods

        #region Helpers
        private void SetIncludeOnlySection(ONNotebook notebook)
        {
            if (notebook==null)
                throw new ArgumentNullException("notebook");

            ONSection section = notebook.GetSection(Config.Current.ExportOnly);
            if (section == null)
            {
                // dont find any section to process
                return;
            }
            notebook.SetNotebookExcludeFlag(true);
            notebook.SetSectionExcludeFlag(section, false);

            // propagate settings
            ONBasedType b = section;
            while (b != notebook)
            {
                if (b is ONSection)
                {
                    (b as ONSection).Exclude = false;
                }
                b = b.Parent;
            }

            Log.Information(string.Format("Export only section [{0}]", section.Name));
            //foreach 
        }

        private void SetExcludeSection(ONNotebook notebook)
        {
            // set sections to exclude
            foreach (string s in Config.Current.ExcludeSections)
            {
                Data.ONSection exSec = notebook.GetSection(s);
                if (exSec != null)
                {
                    Log.Information(string.Format("Exclude section [{0}] from exporting", exSec.Name));
                    exSec.Exclude = true;
                }
            }
        }
        #endregion

        #region PDF Document
        private bool InitDocument(string pathName)
        {
            try
            {
                pdfDocument = new Document(PageSize.A4, 36, 36, 54, 36);
                //pdfDocument = new Document(PageSize.A4);

                pdfWriter = PdfWriter.GetInstance(pdfDocument, new FileStream(pathName, FileMode.Create));
                pdfWriter.SetPdfVersion(PdfWriter.PDF_VERSION_1_5);
                pdfDocument.SetMarginMirroring(true);
                pdfWriter.SetLinearPageMode();
                pdfDocument.Open();
                return true;
            }
            catch (System.Exception ex)
            {
                Log.Error(ex.Message);
                return false;
            }
        }
        private void CloseDocument()
        {
            if (pdfDocument!=null)
                pdfDocument.Close();
            pdfDocument = null;
        }
        #endregion

        #region PDF Combine
        private void PDFCombineAll(string basedName, Data.ONNotebook notebook)
        {
            if (notebook == null)
                throw new ArgumentNullException("notebook");

            if (!InitDocument(Path.Combine(basedName, notebook.Name) + ".pdf"))
                return;

            TOCHandler.Init(pdfDocument, pdfWriter);

            PdfOutline root = pdfWriter.DirectContent.RootOutline;
            TOCHandler.BeginTocEntry();
            foreach (Data.ONSection section in notebook.Sections)
            {
                if (!section.Exclude)
                {
                    TOCHandler.AddTocEntry(section.Name, pdfWriter.CurrentPageNumber);
                    TOCHandler.BeginTocEntry();
                    PDFCombineSection(root, section);
                    TOCHandler.EndTocEntry();
                }
            }
            TOCHandler.EndTocEntry();

            TOCHandler.WriteTocEntries();
            TOCHandler.Close();
            CloseDocument();
        }

        private void PDFMergeAll(string basedName, ONSection sec)
        {
            if (sec == null)
                throw new ArgumentNullException("sec");
            if (sec.Exclude)
            {
                Log.Warning("Trying to export the excluding section. Abort!");
                return;
            }

            if (!InitDocument(Path.Combine(basedName, sec.Name) + ".pdf"))
                return;
            TOCHandler.Init(pdfDocument, pdfWriter);

            TOCHandler.BeginTocEntry();
            PDFCombineSection(pdfWriter.DirectContent.RootOutline, sec);
            TOCHandler.EndTocEntry();

            TOCHandler.WriteTocEntries();
            TOCHandler.Close();
            CloseDocument();
        }
        
        private void PDFCombineSection(PdfOutline parent,
                                        Data.ONSection section)
        {
            if (section == null)
                throw new ArgumentNullException("section");
            if (section.Exclude)
                return;

            PdfOutline gotoPage = null;

            int currentPage = (pdfWriter.CurrentPageNumber == 1) ? pdfWriter.CurrentPageNumber : (pdfWriter.CurrentPageNumber + 1);
            PdfAction action = PdfAction.GotoLocalPage(currentPage, new PdfDestination(PdfDestination.FITH, 806), pdfWriter);
            gotoPage = new PdfOutline(parent, action, section.Name);
            pdfWriter.DirectContent.AddOutline(gotoPage);

            foreach (Data.ONSection subSection in section.SubSections)
            {
                if (!subSection.Exclude)
                {
                    TOCHandler.AddTocEntry(subSection.Name, pdfWriter.CurrentPageNumber);

                    TOCHandler.BeginTocEntry();
                    PDFCombineSection(gotoPage, subSection);
                    TOCHandler.EndTocEntry();
                }
            }
            foreach (Data.ONPage page in section.Pages)
            {
                if (pdfWriter.CurrentPageNumber == 1)
                {
                    TOCHandler.AddTocEntry(page.Name, pdfWriter.CurrentPageNumber);
                }
                else
                {
                    TOCHandler.AddTocEntry(page.Name, pdfWriter.CurrentPageNumber + 1);
                }
                PDFCombinePage(gotoPage, page);
            }
        }

        private void PDFCombinePage(PdfOutline parent,
                                    Data.ONPage page)
        {
            if (page == null)
                throw new ArgumentNullException("page");

            PdfReader reader = new PdfReader(page.PDFFilePath);
            Log.Information(string.Format("Import file {0}", page.PDFFilePath));

            // we retrieve the total number of pages

            int n = reader.NumberOfPages;
            PdfContentByte cb = pdfWriter.DirectContent;
            PdfImportedPage impportPage;
            int rotation;
            int i = 0;
            while (i < n)
            {
                i++;
                pdfDocument.SetPageSize(reader.GetPageSizeWithRotation(i));
                pdfDocument.NewPage();
                impportPage = pdfWriter.GetImportedPage(reader, i);
                rotation = reader.GetPageRotation(i);
                if (rotation == 90 || rotation == 270)
                {
                    cb.AddTemplate(impportPage, 0, -1f, 1f, 0, 0, reader.GetPageSizeWithRotation(i).Height);
                }
                else
                {
                    cb.AddTemplate(impportPage, 1f, 0, 0, 1f, 0, 0);
                }
                if (i == 1)
                {
                    // set outline action at first page
                    PdfAction action = PdfAction.GotoLocalPage(pdfWriter.CurrentPageNumber, new PdfDestination(PdfDestination.FITH, 806), pdfWriter);
                    PdfOutline gotoPage = new PdfOutline(parent, action, page.Name);
                    pdfWriter.DirectContent.AddOutline(gotoPage);
                }
            }
            pdfWriter.FreeReader(reader);
        }
        #endregion

        #region PDF Export
        private void PDFExportSection(string basedName, Data.ONSection section)
        {
            if (section == null)
                throw new ArgumentNullException("section");

            string baseDir = Path.Combine(basedName, section.Name);
            Directory.CreateDirectory(baseDir);
            // now export sub sections
            foreach (Data.ONSection subSection in section.SubSections)
            {
                if (!subSection.Exclude)
                {
                    PDFExportSection(baseDir, subSection);
                }
            }
            // now export pages
            foreach (Data.ONPage page in section.Pages)
            {
                PDFExportPage(baseDir, page);
            }
        }

        private void PDFExportPage(string basedName, Data.ONPage page)
        {
            if (page == null)
                throw new ArgumentNullException("page");

            string fileName = Path.Combine(basedName, page.ID);
            fileName += ".pdf";
            page.PDFFilePath = fileName;

            if (File.Exists(fileName))
            {
                if (File.GetCreationTime(fileName) == page.LastModifiedTime)
                {
                    // this page is not modified since last update
                    return;
                }
                else
                {
                    // delete existing file
                    File.Delete(fileName);
                }
            }
            Log.Verbose(string.Format("Exporting page [{0}] to PDF", page.Name));
            OneNoteApplication.Publish(page.ID, fileName, OneNote.PublishFormat.pfPDF, string.Empty);
            File.SetCreationTime(fileName, page.LastModifiedTime);
            File.SetLastWriteTime(fileName, page.LastModifiedTime);
        }
        #endregion

        #endregion
    }
}

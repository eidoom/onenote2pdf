using System;
using OneNote2PDF.Library;
using OneNote2PDF.Library.Data;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace OneNote2PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            OneNote.Application oneApp;
            // Obtain reference to OneNote application
            try
            {
                oneApp = new OneNote.Application();
            }
            catch (Exception e)
            {
                Log.Error(string.Format("Could not obtain reference to OneNote ({0})", e.Message));
                oneApp = null;
                return;
            }

            string outputXML;
            // get till page level
            oneApp.GetHierarchy(null, OneNote.HierarchyScope.hsPages, out outputXML);
            ONNotebookListing onListing = new ONNotebookListing(outputXML);

            if (Config.Current.ShowHelp || Config.Current.Arguments.Count == 0)
            {
                HelpHandler hh = new HelpHandler();
                Log.Information(hh.Print());
                return;
            }

            if (Config.Current.ListAllNotebook)
            {
                string[] notebookNames = onListing.ListAllNotebook();
                foreach (var nb in notebookNames)
                {
                    Console.WriteLine(nb);
                }
                return;
            }

            if (!string.IsNullOrEmpty(Config.Current.NotebookName))
            {
                Log.Information("Query notebook information");
                ONNotebook notebook = onListing.GetNotebook(Config.Current.NotebookName);
                if (notebook == null)
                {
                    Log.Error("Cannot get desired notebook");
                    return;
                }
                Log.Information(string.Format("Notebook name: {0}", notebook.Name));
                Log.Information(string.Format("Notebook ID: {0}", notebook.ID));
                Log.Information(string.Format("Number section: {0}", notebook.Sections.Count));

                Log.Information("Begin exporting ...");

                Export2PDF export2PDF = new Export2PDF();
                export2PDF.BasedPath = Config.Current.CacheFolder;
                export2PDF.OneNoteApplication = oneApp;

                export2PDF.CreateCacheFolder(notebook);

                if (Config.Current.ExportNotebook)
                {
                    Log.Information(string.Format("Exporting entire notebook [{0}]", notebook.Name));
                    export2PDF.Export(Config.Current.OutputPath, notebook);
                }
                if (!string.IsNullOrEmpty(Config.Current.ExportSection))
                {
                    OneNote2PDF.Library.Data.ONSection section = notebook.GetSection(Config.Current.ExportSection);
                    if (section == null)
                    {
                        Log.Error("Cannot find the specified section");
                    }
                    else
                    {
                        Log.Information(string.Format("Exporting section [{0}] ...", section.Name));
                        export2PDF.Export(Config.Current.OutputPath, section);
                    }
                }
                return;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using OneNote2PDF.Library.Data;

namespace OneNote2PDF.Library
{
    class TOCHandler
    {
        #region Private const
        private const string TABLEOFCONTENTS = "Table of Contents";
        private const int MAXTOCLEVEL = 20;
        #endregion

        #region Private variables
        private Document pdfDocument { get; set; }
        private PdfWriter pdfWriter { get; set; }
        private List<TOCEntry> TocEntries { get; set; }
        private int[] TOCNumbering = new int[MAXTOCLEVEL];
        private int currentLevel;
        #endregion
        
        #region Public methods
        public void Init(Document document, PdfWriter writer)
        {
            pdfDocument = document;
            pdfWriter = writer;

            ClearNumbering();
            TocEntries = new List<TOCEntry>();
            currentLevel = 0;
        }
        public void Close()
        {
            ClearNumbering();
            TocEntries.Clear();
        }

        public void BeginTocEntry()
        {
            if (currentLevel < 0 || currentLevel >= MAXTOCLEVEL)
                throw new ArgumentOutOfRangeException("level", string.Format("Level must be in the range [1..{0}]", MAXTOCLEVEL));

            if (TocEntries == null)
            {
                Log.Error("Init must be called before calling EndTocEntry");
                throw new InvalidOperationException("Init must be called before calling EndTocEntry");
            }
            // increase current level
            ++currentLevel;
            TOCNumbering[currentLevel] = 1;
        }
        public void EndTocEntry()
        {
            if (currentLevel <= 0 || currentLevel >= MAXTOCLEVEL)
                throw new ArgumentOutOfRangeException("level", string.Format("Level must be in the range [1..{0}]", MAXTOCLEVEL));
            if (TocEntries == null)
            {
                Log.Error("Init must be called before calling EndTocEntry");
                throw new InvalidOperationException("Init must be called before calling EndTocEntry");
            }
            for (int i = currentLevel; i < MAXTOCLEVEL; ++i)
            {
                TOCNumbering[i] = -1;
            }
            --currentLevel;
        }
        public void AddTocEntry(string title, int pageNumber)
        {
            if (TocEntries == null)
            {
                Log.Error("Init() must be called before calling AddTocEntry()");
                throw new InvalidOperationException("Init() must be called before calling AddTocEntry()");
            }
            if (currentLevel<=0)
            {
                Log.Error("BeginTocEntry must be called before calling AddTocEntry()");
                throw new InvalidOperationException("BeginTocEntry() must be called before calling AddTocEntry()");
            }

            TocEntries.Add(new TOCEntry() { Title = title, Level = currentLevel, PageNumber = pageNumber, Numbering = AutoTocNumbering(currentLevel) });
        }
        public void WriteTocEntries()
        {
            if (!Config.Current.ShowTOC)
            {
                return;
            }
            if (pdfWriter == null || pdfDocument == null)
            {
                Log.Error("Uninitialized property pdfWriter or pdfDocument. Pass valid values and try again");
                return;
            }
            Log.Information("Writing Table of Contents ...");
            int beforeToc = pdfWriter.CurrentPageNumber;

            pdfWriter.PageEvent = new TocFooterHandler() { StartPageIndex = beforeToc, TOCString = TABLEOFCONTENTS };
            pdfDocument.NewPage();

            PdfAction action = PdfAction.GotoLocalPage(pdfWriter.CurrentPageNumber, new PdfDestination(PdfDestination.FITH, 806), pdfWriter);
            PdfOutline gotoPage = new PdfOutline(pdfWriter.RootOutline, action, TABLEOFCONTENTS);
            pdfWriter.DirectContent.AddOutline(gotoPage);

            Paragraph ToCTitle = new Paragraph(TABLEOFCONTENTS, new Font(BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.WINANSI, true), 20));
            ToCTitle.SpacingAfter = 20;
            pdfDocument.Add(ToCTitle);

            Font tocFont = new Font(BaseFont.CreateFont(Path.Combine(Environment.SystemDirectory, "../fonts/arial.ttf"), BaseFont.IDENTITY_H, true), 12);

            foreach (TOCEntry tocEntry in TocEntries)
            {
                // Auto advance to next page
                float currentY = pdfWriter.GetVerticalPosition(false);
                if (currentY < 90)
                    pdfDocument.NewPage();

                // write toc entry title
                Paragraph a = new Paragraph();
                Chunk c;
                if (tocEntry.Level <= Config.Current.TOCLevel)
                    c = new Chunk(string.Format("{0} {1}", tocEntry.Numbering, tocEntry.Title), tocFont);
                else
                    c = new Chunk(tocEntry.Title, tocFont);

                int destinationPage = tocEntry.PageNumber;
                PdfAction gotoDestination = PdfAction.GotoLocalPage(destinationPage, new PdfDestination(PdfDestination.FITH, 806), pdfWriter);
                c.SetAction(gotoDestination);
                a.Add(c);
                a.IndentationLeft = (tocEntry.Level - 1) * 20 + 60;
                a.FirstLineIndent = -40;
                a.IndentationRight = 30;

                pdfDocument.Add(a);

                // write toc entry page number
                currentY = pdfWriter.GetVerticalPosition(false);
                PdfContentByte cb = pdfWriter.DirectContent;
                cb.SaveState();
                cb.BeginText();
                BaseFont bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
                cb.SetFontAndSize(bf, 12);
                cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, destinationPage.ToString(), pdfDocument.Right, currentY, 0);
                cb.EndText();
                cb.RestoreState();
            }
            Log.Information("Reordering pages ...");
            int totalPages = pdfWriter.CurrentPageNumber;

            // now reorder the pages
            int[] reorder = new int[totalPages];
            for (int i = 0; i < totalPages; i++)
            {
                reorder[i] = i + beforeToc + 1;
                if (reorder[i] > totalPages)
                    reorder[i] -= totalPages;
            }
            pdfDocument.NewPage();
            pdfWriter.PageEvent = null;

            pdfWriter.ReorderPages(reorder);

            Log.Information(string.Format("Number of pages for TOC is {0}", totalPages - beforeToc));
            Log.Verbose("Set page labels");
            PdfPageLabels pageLabels = new PdfPageLabels();
            pageLabels.AddPageLabel(1, PdfPageLabels.LOWERCASE_ROMAN_NUMERALS);
            pageLabels.AddPageLabel(totalPages - beforeToc + 1, PdfPageLabels.DECIMAL_ARABIC_NUMERALS);
            pdfWriter.PageLabels = pageLabels;
        }
        #endregion

        #region Private members
        private void ClearNumbering()
        {
            for (int i = 0; i < MAXTOCLEVEL; ++i)
            {
                TOCNumbering[i] = -1;
            }
        }

        private string AutoTocNumbering(int level)
        {
            StringBuilder numbering = new StringBuilder();
            for (int i = 1; i < level; ++i)
            {
                numbering.AppendFormat("{0}.", TOCNumbering[i] - 1);
            }
            numbering.AppendFormat("{0}", TOCNumbering[level]);
            TOCNumbering[level]++;
            return numbering.ToString();
        }
        #endregion
    }
}

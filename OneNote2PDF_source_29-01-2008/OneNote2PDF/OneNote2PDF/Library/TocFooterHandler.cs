using System;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace OneNote2PDF.Library
{
    class Converter
    {
        public static string NumberToRoman(int number, bool lowerCase)
        {
            // Set up key numerals and numeral pairs
            int[] values = new int[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            string[] numerals = new string[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };

            // Validate
            if (number < 0 || number > 3999)
                throw new ArgumentException("Value must be in the range 0 - 3,999.");
            
            if (number == 0)
                return "N";

            // Initialise the string builder
            StringBuilder result = new StringBuilder();
            // Loop through each of the values to diminish the number
            for (int i = 0; i < 13; i++)
            {
                // If the number being converted is less than the test value, append
                // the corresponding numeral or numeral pair to the resultant string
                while (number >= values[i])
                {
                    number -= values[i];
                    result.Append(numerals[i]);
                }
            }
            // Done
            if (lowerCase)
                return result.ToString().ToLower();
            else
                return result.ToString();
        }
    }
    /// <summary>
    /// PdfPageEvent handler to add header to Table-of-Contents pages
    /// </summary>
    class TocFooterHandler : IPdfPageEvent
    {
        private const int TOCINDENT = 40; 
        private const int TOCTOPOFFSET = 20;
        private const int TOCFONTSIZE = 12;

        /// <summary>
        /// The font that will be used for Toc string.
        /// </summary>
        protected BaseFont helv;
        /// <summary>
        /// The font that will be used for TOC Number.
        /// </summary>
        protected BaseFont helvb;
        public int StartPageIndex { get; set; }
        public string TOCString { get; set; }

        /// <summary>
        /// Constructs an Event that adds a Header and a Footer.
        /// </summary>
        public TocFooterHandler()
        {
            helv = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI,
                    BaseFont.NOT_EMBEDDED);
            helvb = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.WINANSI,
                    BaseFont.NOT_EMBEDDED);
        }

        void IPdfPageEvent.OnEndPage(PdfWriter writer, Document document)
        {
            int pageNumber = writer.PageNumber - StartPageIndex;
            if (pageNumber == 1)
            {
                // skip first page
                return;
            }

            PdfContentByte cb = writer.DirectContent;
            cb.SaveState();

            String text = Converter.NumberToRoman(pageNumber, true);
            float textBase = document.Top + TOCTOPOFFSET;
            float textSize = helvb.GetWidthPoint(text, TOCFONTSIZE);
            cb.BeginText();
            
            // center footer
            //float adjust = helv.GetWidthPoint("0", 12);
            //cb.SetTextMatrix((document.Left + document.Right)/2 - textSize/2 - adjust, textBase);
            //cb.ShowText(text);
            //cb.EndText();

            if ((pageNumber % 2) == 1)
            {
                // for odd numbers, show the header at the left
                cb.SetFontAndSize(helvb, TOCFONTSIZE);
                cb.SetTextMatrix(document.Left, textBase);
                cb.ShowText(text);

                cb.SetFontAndSize(helv, TOCFONTSIZE);
                cb.SetTextMatrix(document.Left + TOCINDENT, textBase);
                cb.ShowText(TOCString);
                cb.EndText();
            }
            else
            {
                // for even numbers, show the header at the right
                float textTOCSize = helv.GetWidthPoint(TOCString, TOCFONTSIZE);

                cb.SetFontAndSize(helv, TOCFONTSIZE);
                cb.SetTextMatrix(document.Right - textTOCSize - TOCINDENT, textBase);
                cb.ShowText(TOCString);

                cb.SetFontAndSize(helvb, TOCFONTSIZE);
                cb.SetTextMatrix(document.Right - textSize, textBase);
                cb.ShowText(text);
                cb.EndText();
            }

            cb.RestoreState();
        }

        #region Unimplemented IPdfPageEvent Members
        void IPdfPageEvent.OnStartPage(PdfWriter writer, Document document)
        {
        }

        void IPdfPageEvent.OnChapter(PdfWriter writer, Document document, float paragraphPosition, Paragraph title)
        {
        }

        void IPdfPageEvent.OnChapterEnd(PdfWriter writer, Document document, float paragraphPosition)
        {
        }

        void IPdfPageEvent.OnCloseDocument(PdfWriter writer, Document document)
        {
        }

        void IPdfPageEvent.OnGenericTag(PdfWriter writer, Document document, Rectangle rect, string text)
        {
        }

        void IPdfPageEvent.OnOpenDocument(PdfWriter writer, Document document)
        {
        }

        void IPdfPageEvent.OnParagraph(PdfWriter writer, Document document, float paragraphPosition)
        {
        }

        void IPdfPageEvent.OnParagraphEnd(PdfWriter writer, Document document, float paragraphPosition)
        {
        }

        void IPdfPageEvent.OnSection(PdfWriter writer, Document document, float paragraphPosition, int depth, Paragraph title)
        {
        }

        void IPdfPageEvent.OnSectionEnd(PdfWriter writer, Document document, float paragraphPosition)
        {
        }

        #endregion
    }
}

using System.Collections.Generic;

namespace OneNote2PDF.Library.Data
{
    public class ONSection : ONBasedType
    {
        public ONSection()
        {
            Pages = new List<ONPage>();
            SubSections = new List<ONSection>();
            // default is false 
            Exclude = false;
        }
        public List<ONPage> Pages { get; set; }
        public List<ONSection> SubSections { get; set; }

        public bool Exclude { get; set; }
    }
}

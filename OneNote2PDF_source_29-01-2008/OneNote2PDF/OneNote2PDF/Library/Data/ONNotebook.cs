using System.Linq;
using System.Collections.Generic;
using System;

namespace OneNote2PDF.Library.Data
{
    public class ONNotebook: ONBasedType
    {
        public List<ONSection> Sections { get; set; }

        public void SetNotebookExcludeFlag(bool isExclude)
        {
            foreach (ONSection nbSection in Sections)
            {
                SetSectionExcludeFlag(nbSection, isExclude);
            }
        }

        public void SetSectionExcludeFlag(ONSection section, bool isExclude)
        {
            if (section == null)
                return;

            Queue<ONSection> queue = new Queue<ONSection>();

            queue.Enqueue(section);
            while (queue.Count > 0)
            {
                ONSection s = queue.Dequeue();
                s.Exclude = isExclude;
                foreach (ONSection ss in s.SubSections)
                {
                    queue.Enqueue(ss);
                }
            }
        }

        public ONSection GetSection(string sectionName)
        {
            string realSName;
            string extra;

            SplitPath(sectionName, out realSName, out extra);

            try
            {
                ONSection result = (from section in Sections
                                    where realSName.ToLower() == section.Name.ToLower()
                                    select section).First();

                if (string.IsNullOrEmpty(extra))
                    return result;
                else
                    return GetSubSection(result, extra);
            }
            catch (InvalidOperationException)
            {
                // The source sequence is empty.
                return null;
            }
            catch (ArgumentNullException)
            {
                return null;
            }

        }
        private ONSection GetSubSection(ONSection section, string sectionName)
        {
            string realSName;
            string extra;

            SplitPath(sectionName, out realSName, out extra);

            try
            {
                ONSection results = (from subSection in section.SubSections
                                     where realSName.ToLower() == subSection.Name.ToLower()
                                     select subSection).First();

                if (string.IsNullOrEmpty(extra))
                    return results;
                else
                    return GetSubSection(results, extra);
            }
            catch (InvalidOperationException)
            {
                // The source sequence is empty.
                return null;
            }
            catch (ArgumentNullException)
            {
                return null;
            }
        }
        private void SplitPath(string path, out string first, out string extra)
        {
            first = string.Empty;
            extra = string.Empty;

            int nPos = path.IndexOf('/');
            if (nPos == -1)
            {
                first = path;
                extra = string.Empty;
            }
            else
            {
                first = path.Substring(0, nPos);
                extra = path.Remove(0, nPos + 1);
            }
        }
    }
}

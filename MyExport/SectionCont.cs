using System;
using System.Collections.Generic;
using System.Text;

namespace MyExportLibrary
{

    class SectionInfo
    {
        public string FileName;
        public int StartInd;
    }
    ///////////////////////////////////////////////////////////////////////////////

    class Sections
    {
        private List<SectionInfo> SectionsContainer;
        ///////////////////////////////////////////////////////////////////////////////
        public Sections()
        {
            SectionsContainer = new List<SectionInfo>();
        }
        ///////////////////////////////////////////////////////////////////////////////
        public int Count
        {
            get
            {
                return SectionsContainer.Count;
            }
        }
        ///////////////////////////////////////////////////////////////////////////////
        public void Add(string FileName, int StartInd)
        {
            SectionInfo Temp = new SectionInfo();
            Temp.FileName = FileName;
            Temp.StartInd = StartInd;
            SectionsContainer.Add(Temp);
        }
        ///////////////////////////////////////////////////////////////////////////////
        public SectionInfo this[int i]
        {
            get
            {
                return SectionsContainer[i];
            }
        }
    }
}

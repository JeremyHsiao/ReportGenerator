using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReportApplication
{
    class NotReportFileRecord
    {
        private String path;
        private String filename;
        private Boolean openfileOK;
        private Boolean findWorksheetOK;
        private Boolean findAnyKeyword;
        private Boolean otherFailure;

        //public String Path   // property
        //{
        //    get { return path; }   // get method
        //    set { path = value; }  // set method
        //}
        //public String Filename   // property
        //{
        //    get { return filename; }   // get method
        //    set { filename = value; }  // set method
        //}
        //public Boolean FilenameOK   // property
        //{
        //    get { return filenameOK; }   // get method
        //    set { filenameOK = value; }  // set method
        //}
        //public Boolean SheetnameOK   // property
        //{
        //    get { return sheetnameOK; }   // get method
        //    set { sheetnameOK = value; }  // set method
        //}
        //public Boolean ItemOK   // property
        //{
        //    get { return itemOK; }   // get method
        //    set { itemOK = value; }  // set method
        //}
        //public Boolean CaptionOK   // property
        //{
        //    get { return captionOK; }   // get method
        //    set { captionOK = value; }  // set method
        //}
        //public Boolean OtherFailure   // property
        //{
        //    get { return otherFailure; }   // get method
        //    set { otherFailure = value; }  // set method
        //}

        //
        public NotReportFileRecord() { this.path = this.filename = ""; this.openfileOK = false; this.otherFailure = false; }
        public NotReportFileRecord(String path="", String filename="") 
        {
            this.path = path;
            this.filename = filename;
            this.openfileOK = false; 
            this.otherFailure = false; 
        }
        //public NotReportFileRecord(String path, String filename, Boolean filenameOK,
        //                            Boolean sheetnameOK, Boolean itemOK, Boolean captionOK, Boolean otherFailure=false)
        //{ SetRecord(path, filename, filenameOK, sheetnameOK, itemOK, captionOK, otherFailure); }

        // only set fail flag, don't change if it doesn't fail
        public void SetFlagFail(Boolean openfileFail = false, Boolean findWorksheetFail = false, Boolean findNoKeyword = false, 
                                Boolean otherFailure = false)
        {
            if (openfileFail) { this.openfileOK = false; }
            if (findWorksheetFail) { this.findWorksheetOK = false; }
            if (findNoKeyword) { this.findAnyKeyword = false; }
            if (otherFailure) { this.otherFailure = true; }
        }
        // only set OK flag, don't change if it doesn't OK
        public void SetFlagOK(Boolean openfileOK = false, Boolean findWorksheetOK = false, Boolean findAnyKeyword = false, 
                              Boolean otherAllOK = false)
        {
            if (openfileOK) { this.openfileOK = true; }
            if (findWorksheetOK) { this.findWorksheetOK = true; }
            if (findAnyKeyword) { this.findAnyKeyword = true; }
            if (otherAllOK) { this.otherFailure = false; }
        }
        public void GetFlagValue(out Boolean openfileOK, out Boolean findWorksheetOK, out Boolean findAnyKeyword, 
                                out Boolean otherFailure)
        {
            openfileOK = this.openfileOK;
            findWorksheetOK = this.findWorksheetOK;
            findAnyKeyword = this.findAnyKeyword;
            otherFailure = this.otherFailure;
        }
        public void SetFlagValue(Boolean openfileOK, Boolean findWorksheetOK, Boolean findAnyKeyword, Boolean otherFailure = false)
        {
            this.openfileOK = openfileOK;
            this.findWorksheetOK = findWorksheetOK;
            this.findAnyKeyword = findAnyKeyword;
            this.otherFailure = otherFailure;
        }
        public void GetRecord(out String path, out String filename, out Boolean openfileOK, out Boolean findWorksheetOK, 
                            out Boolean findAnyKeyword, out Boolean otherFailure)
        {
            path = this.path;
            filename = this.filename;
            this.GetFlagValue(out openfileOK, out findWorksheetOK, out findAnyKeyword, out otherFailure);
        }

        public void SetRecord(String path, String filename, Boolean openfileOK, Boolean findWorksheetOK, Boolean findAnyKeyword, 
                                Boolean otherFailure = false)
        {
            this.path = path;
            this.filename = filename;
            this.SetFlagValue(openfileOK, findWorksheetOK, findAnyKeyword, otherFailure);
        }
    }

    class KeyWordListReport
    {
        static private string Template_Excel = "Template_Excel";
        static private string WS_KeyWord_List = "Keyword_List";
        static private string WS_NotKeyWord_List = "Not_Keyword_File";
        static private string Output_Excel = "Output_Excel";

        //static public 
    }
}

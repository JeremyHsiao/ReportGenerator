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
        private Boolean filenameOK;
        private Boolean sheetnameOK;
        private Boolean itemOK;
        private Boolean captionOK;
        //private Boolean OtherCaptiponOK;
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
        public NotReportFileRecord() { this.otherFailure=false; }
        public NotReportFileRecord(String path, String filename) { this.path = path; this.filename = filename; this.otherFailure=false; }
        public NotReportFileRecord(String path, String filename, Boolean filenameOK,
                                    Boolean sheetnameOK, Boolean itemOK, Boolean captionOK, Boolean otherFailure=false)
        { SetRecord(path, filename, filenameOK, sheetnameOK, itemOK, captionOK, otherFailure); }

        public void SetFlagFail(Boolean filenamefail = false, Boolean sheetnamefail = false, Boolean itemfail = false, 
                                Boolean captionfail = false, Boolean otherFailure = false)
        {
            if (filenamefail) { this.filenameOK = false; }
            if (sheetnamefail) { this.sheetnameOK = false; }
            if (itemfail) { this.itemOK = false; }
            if (captionfail) { this.captionOK = false; }
            if (otherFailure) { this.otherFailure = true; }
        }
        public void SetFlagOK(Boolean filenameOK = false, Boolean sheetnameOK = false, Boolean itemOK = false, 
                                Boolean captionOK = false, Boolean otherOK = false)
        {
            if (filenameOK) { this.filenameOK = true; }
            if (sheetnameOK) { this.sheetnameOK = true; }
            if (itemOK) { this.itemOK = true; }
            if (captionOK) { this.captionOK = true; }
            if (otherOK) { this.otherFailure = false; }
        }
        public void GetFlagValue(out Boolean filenameOK, out Boolean sheetnameOK, out Boolean itemOK, 
                                out Boolean captionOK, out Boolean otherFailure)
        {
            filenameOK = this.filenameOK;
            sheetnameOK = this.sheetnameOK;
            itemOK = this.itemOK;
            captionOK = this.captionOK;
            otherFailure = this.otherFailure;
        }
        public void SetFlagValue(Boolean filenameOK, Boolean sheetnameOK, Boolean itemOK, Boolean captionOK, Boolean otherFailure = false)
        {
            this.filenameOK = filenameOK;
            this.sheetnameOK = sheetnameOK;
            this.itemOK = itemOK;
            this.captionOK = captionOK;
            this.otherFailure = otherFailure;
        }
        public void GetRecord(out String path, out String filename, out Boolean filenameOK, out Boolean sheetnameOK, 
                            out Boolean itemOK, out Boolean captionOK, out Boolean otherFailure)
        {
            path = this.path;
            filename = this.filename;
            filenameOK = this.filenameOK;
            sheetnameOK = this.sheetnameOK;
            itemOK = this.itemOK;
            captionOK = this.captionOK;
            otherFailure = this.otherFailure;
        }

        public void SetRecord(String path, String filename, Boolean filenameOK, Boolean sheetnameOK, Boolean itemOK, 
                                Boolean captionOK, Boolean otherFailure)
        {
            this.path = path;
            this.filename = filename;
            this.filenameOK = filenameOK;
            this.sheetnameOK = sheetnameOK;
            this.itemOK = itemOK;
            this.captionOK = captionOK;
            this.otherFailure = otherFailure;
        }
    }

    class KeyWordListReport
    {
        static private string Template_Excel = "Template_Excel";
        static private string WS_KeyWord_List = "Keyword_List";
        static private string WS_NotKeyWord_List = "Not_Keyword_File";
        static private string  Output_Excel = "Output_Excel";

        //static public 
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReportApplication
{
    public class CopyTestReport
    {
        public String source_path;
        public String source_folder;
        public String source_group;
        public String source_filename;
        public String destination_path;
        public String destination_folder;
        public String destination_group;
        public String destination_filename;
        public String destination_assignee;

        public String Get_SRC_Directory()
        {
            String path = this.source_path;
            if (this.source_folder != "")
            {
                path = Storage.CominePath(path, this.source_folder);
            }
            if (this.source_group != "")
            {
                path = Storage.CominePath(path, this.source_group);
            }
            return path;
        }

        public String Get_DEST_Directory()
        {
            String path = this.destination_path;
            if (this.destination_folder != "")
            {
                path = Storage.CominePath(path, this.destination_folder);
            }
            if (this.destination_group != "")
            {
                path = Storage.CominePath(path, this.destination_group);
            }
            return path;
        }

        public String Get_SRC_FullFilePath()
        {
            String path = this.Get_SRC_Directory();
            String file = this.source_filename + ".xlsx";
            String fullfilepath = Storage.GetValidFullFilename(path, file);
            return fullfilepath;
        }

        public String Get_DEST_FullFilePath()
        {
            String path = this.Get_DEST_Directory();
            String file = this.destination_filename + ".xlsx";
            String fullfilepath = Storage.GetValidFullFilename(path, file);
            return fullfilepath;
        }

    }

    //public class ReportMapping
    //{
    //                //ctp.source_folder = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
    //                //ctp.source_group = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
    //                //ctp.source_filename = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
    //                //ctp.destination_path = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
    //                //ctp.destination_folder = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
    //                //ctp.destination_group = ExcelAction.GetCellTrimmedString(ws, row_index, col_index++);
    //                //ctp.destination_filename = ExcelAction.GetCellTrimmedString(ws, row_index, col_index);
    //    private enum Name
    //    {
    //        source_path = 0,
    //        source_folder,
    //        source_group,
    //        source_filename,
    //        destination_path,
    //        destination_folder,
    //        destination_group,
    //        destination_filename,
    //        destination_assignee,
    //    }

    //    private static int EnumNameCount = Enum.GetNames(typeof(Name)).Length;

    //    private List<String> list_of_string;

    //    private void InitReportMapping()
    //    {
    //        list_of_string = new List<String>();
    //        for (int index = 0; index < EnumNameCount; index++)
    //        {
    //            list_of_string.Add("");
    //        }
    //    }

    //    // Note: here the sequence of member is pre-defined as source(path/folder/group/report), destination(path/folder/group/report/assignee)
    //    private void SetupReportMapping(List<String> member)
    //    {
    //        if (member.Count >= EnumNameCount)
    //        {
    //            int index = 0;
    //            do
    //            {
    //                list_of_string[index] = member[index];
    //            }
    //            while (++index < member.Count);
    //        }
    //        else
    //        {
    //            int index = 0;
    //            do
    //            {
    //                list_of_string[index] = member[index];
    //            }
    //            while (++index < member.Count);
    //            do
    //            {
    //                list_of_string[index] = "";
    //            }
    //            while (++index < EnumNameCount);
    //        }
    //    }

    //    public ReportMapping() { InitReportMapping(); }

    //    public ReportMapping(List<String> member) { InitReportMapping(); SetupReportMapping(member); }

    //    public String Source_Path   // property
    //    {
    //        get { return list_of_string[(int)Name.source_path]; }   // get method
    //        set { list_of_string[(int)Name.source_path] = value; }  // set method
    //    }

    //    public String Source_Folder   // property
    //    {
    //        get { return list_of_string[(int)Name.source_folder]; }   // get method
    //        set { list_of_string[(int)Name.source_folder] = value; }  // set method
    //    }

    //    public String Source_Group   // property
    //    {
    //        get { return list_of_string[(int)Name.source_group]; }   // get method
    //        set { list_of_string[(int)Name.source_group] = value; }  // set method
    //    }

    //    public String Source_Report   // property
    //    {
    //        get { return list_of_string[(int)Name.source_filename]; }   // get method
    //        set { list_of_string[(int)Name.source_filename] = value; }  // set method
    //    }

    //    public String Destination_Path   // property
    //    {
    //        get { return list_of_string[(int)Name.destination_path]; }   // get method
    //        set { list_of_string[(int)Name.destination_path] = value; }  // set method
    //    }

    //    public String Destination_Folder   // property
    //    {
    //        get { return list_of_string[(int)Name.destination_folder]; }   // get method
    //        set { list_of_string[(int)Name.destination_folder] = value; }  // set method
    //    }

    //    public String Destination_Group   // property
    //    {
    //        get { return list_of_string[(int)Name.destination_group]; }   // get method
    //        set { list_of_string[(int)Name.destination_group] = value; }  // set method
    //    }

    //    public String Destination_Report   // property
    //    {
    //        get { return list_of_string[(int)Name.destination_filename]; }   // get method
    //        set { list_of_string[(int)Name.destination_filename] = value; }  // set method
    //    }

    //    public String Destination_Assignee   // property
    //    {
    //        get { return list_of_string[(int)Name.destination_assignee]; }   // get method
    //        set { list_of_string[(int)Name.destination_assignee] = value; }  // set method
    //    }

    //}


}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Globalization;

namespace ExcelReportApplication
{
    class FileFunction
    {
        static public String UsesrSelectFilename()
        {
            String ret_str = "";
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select file";
            dialog.InitialDirectory = Directory.GetCurrentDirectory();
            dialog.Filter = "Excel files (*.xls/xlsx)|*.xls;*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                ret_str = dialog.FileName;
                // It seems that File-existing check is done after "Select" button is pressed
                // So no need to check here.
                /*
                if (!CheckFileExist(ret_str))
                {
                    MessageBox.Show("Selected file does not exist so filename remains unchanged.\n");
                    ret_str = "";
                }
                */
            }
            return ret_str;
        }

        static public bool CheckFileExist(String Filename)
        {
            bool ret;
            if (!File.Exists(Filename))
            {
                ret = false;
            }
            else
            {
                ret = true;
            }
            return ret;
        }

        static public String GetFullFilePath(String Filename)
        {
            String ret ="";
            try
            {
                ret = Path.GetFullPath(Filename);
            }
            catch
            {
                // "" will be returned if exceptions
            }
            return ret;
        }
    }
}

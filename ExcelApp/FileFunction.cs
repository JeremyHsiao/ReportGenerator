using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace ExcelReportApplication
{
    class FileFunction
    {
        // Possible multiple-directories selection, so return String[]
        static public String[] UsersSelectDirectory(String init_dir =@".\", bool multiple = true, String title = "Select Folder(s)")
        {
            var openFolder = new CommonOpenFileDialog();
            openFolder.AllowNonFileSystemItems = true;
            openFolder.Multiselect = multiple;                 
            openFolder.IsFolderPicker = true;
            openFolder.Title = title;

            String default_dir = GetDirectoryName(GetFullPath(init_dir));
            if (DirectoryExists(default_dir))
            {
                openFolder.InitialDirectory = default_dir;
            }
            else
            {
                openFolder.InitialDirectory = GetCurrentDirectory();
            }

            if (openFolder.ShowDialog() == CommonFileDialogResult.Ok)
            {
                // get all the directories in selected dirctory
                var dirs = openFolder.FileNames.ToArray();
                return dirs;
            }
            else
            {
                String[] ret_empty = new String[1];
                ret_empty[0] = "";
                return ret_empty;
            }

        }

        // Sigle-directory selection, so return just String
        static public String UsersSelectDirectory()
        {
            return UsersSelectDirectory(init_dir:GetCurrentDirectory());
        }
        static public String UsersSelectDirectory(String init_dir)
        {
            return UsersSelectDirectory(title: "Select a folder", init_dir: init_dir, multiple: false)[0];
        }


        // Possible multiple-file selection, so return String[]
        static public String[] UsesrSelectFilename(String init_dir = @".\", bool multiple = true, String title = "Select File(s)")
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = multiple;
            dialog.Title = title;

            String default_dir = GetDirectoryName(GetFullPath(init_dir));
            if (DirectoryExists(default_dir))
            {
                dialog.InitialDirectory = default_dir;
            }
            else
            {
                dialog.InitialDirectory = GetCurrentDirectory();
            }

            dialog.Filter = "Excel files (*.xls/xlsx)|*.xls;*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                var files = dialog.FileNames.ToArray();
                return files;
            }
            else
            {
                String[] ret_empty = new String[1];
                ret_empty[0] = "";
                return ret_empty;
            }
        }

        // Sigle-file selection, so return just String
        static public String UsesrSelectFilename()
        {
            return UsesrSelectFilename(init_dir: GetCurrentDirectory());
        }
        static public String UsesrSelectFilename(String init_dir)
        {
            return UsesrSelectFilename(title: "Select a file", init_dir: init_dir, multiple: false)[0];
        }

        static public bool FileExists(String Filename)
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

        static public bool DirectoryExists(String dir)
        {
            bool ret;
            if (!Directory.Exists(dir))
            {
                ret = false;
            }
            else
            {
                ret = true;
            }
            return ret;
        }

        static public String GetDirectoryName(String Filename)
        {
            String ret = "";
            try
            {
                ret = Path.GetDirectoryName(Filename);
            }
            catch
            {
                // "" will be returned if exceptions
            }
            return ret;
        }

        static public String GetFullPath(String Filename)
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

        static public String GetFileNameWithoutExtension(String Filename)
        {
            String ret = "";
            try
            {
                ret = Path.GetFileNameWithoutExtension(Filename);
            }
            catch
            {
                // "" will be returned if exceptions
            }
            return ret;
        }

        static public String GetExtension(String Filename)
        {
            String ret = "";
            try
            {
                ret = Path.GetExtension(Filename);
            }
            catch
            {
                // "" will be returned if exceptions
            }
            return ret;
        }

        static public String GetCurrentDirectory()
        {
            String ret = "";
            try
            {
                ret = Directory.GetCurrentDirectory();
            }
            catch
            {
                // "" will be returned if exceptions
            }
            return ret;
        }

        // ();

        static public String GenerateFilenameWithDateTime(String Filename, String new_ext = "")
        {
            String ret = "";
            try
            {
                // Save as another file //yyyyMMddHHmmss
                string path, name, dt, ext;

                path = Path.GetDirectoryName(Filename);             // path without '\'
                name = Path.GetFileNameWithoutExtension(Filename);  // filename only without path
                dt = DateTime.Now.ToString("yyyyMMddHHmmss");       // ex: 20220801160000
                if (new_ext == "")
                {
                    ext = Path.GetExtension(Filename);                  // extension with '.' 
                }
                else
                {
                    ext = new_ext;
                }
                ret = path + @"\" + name + "_" + dt + ext;

                // Filename ==  path + @"\" + name            + ext
                // ret      ==  path + @"\" + name + "_" + dt + ext;
            }
            catch
            {
                // "" will be returned if exceptions
            }
            return ret;
        }

        static public String GenerateDirectoryNameWithDateTime(String dir)
        {
            String ret = "";
            try
            {
                // directory name adding "_yyyyMMddHHmmss"
                string path, dt;

                if (dir[dir.Length - 1] == '\\')
                {
                    dir.Substring(0, dir.Length - 1); // remove '\' at the end
                }
                path = dir;                     
                dt = DateTime.Now.ToString("yyyyMMddHHmmss");       // ex: 20220801160000
                ret = path + "_" + dt;

                // dir  ==  path + @"\"(?)
                // ret  ==  path + "_" + dt;
            }
            catch
            {
                // "" will be returned if exceptions
            }
            return ret;
        }
    }
}

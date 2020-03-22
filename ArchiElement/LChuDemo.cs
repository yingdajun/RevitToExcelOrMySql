using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Windows.Forms;
namespace ArchiElement
{
    class LChuDemo
    {
        public static void dataTableToCsv(System.Data.DataTable table, string file)
        {
            string title = "";
            FileStream fs = new FileStream(file, FileMode.OpenOrCreate);
            StreamWriter sw = new StreamWriter(new BufferedStream(fs), System.Text.Encoding.Default);
            for (int i = 0; i < table.Columns.Count; i++)
            {
                title += table.Columns[i].ColumnName + "\t";
            }
            title = title.Substring(0, title.Length - 1) + "\n";
            sw.Write(title);
            foreach (DataRow row in table.Rows)
            {
                string line = "";
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    line += row[i].ToString().Trim() + "\t";
                }
                line = line.Substring(0, line.Length - 1) + "\n";
                sw.Write(line);
            }
            sw.Close();
            fs.Close();
            TaskDialog.Show("使用提示", "更新数据前请把当前的EXCEL关闭");
        }
        public static bool CopyOldLabFilesToNewLab(string sourcePath, string savePath)
        {
            if (!Directory.Exists(savePath))
            {
                Directory.CreateDirectory(savePath);
            }
            #region           
            try
            {
                string[] labDirs = Directory.GetDirectories(sourcePath); string[] labFiles = Directory.GetFiles(sourcePath); if (labFiles.Length > 0)
                {
                    for (int i = 0; i < labFiles.Length; i++)
                    {
                        if (Path.GetExtension(labFiles[i]) != ".lab")
                        {
                            File.Copy(sourcePath + "\\" + Path.GetFileName(labFiles[i]), savePath + "\\" + Path.GetFileName(labFiles[i]), true);
                        }
                    }
                }
                if (labDirs.Length > 0)
                {
                    for (int j = 0; j < labDirs.Length; j++)
                    {
                        Directory.GetDirectories(sourcePath + "\\" + Path.GetFileName(labDirs[j]));
                        CopyOldLabFilesToNewLab(sourcePath + "\\" + Path.GetFileName(labDirs[j]), savePath + "\\" + Path.GetFileName(labDirs[j]));
                    }
                }
            } 
            catch (Exception)
            {
                return false;
            }
            #endregion
            return true;
        }
        public static string PickFolderInfo(string excel)
        {
            FolderBrowserDialog diaglog = new FolderBrowserDialog();
            string sourcePath = null;
            if (diaglog.ShowDialog() == DialogResult.OK)
            {
                sourcePath = diaglog.SelectedPath;
            }
            TaskDialog.Show("表格提取", "提取的EXCEL已经放置到" + sourcePath.ToString() + "下面请查看");
            CopyOldLabFilesToNewLab(excel, sourcePath);
            return sourcePath;
        }
    }
}

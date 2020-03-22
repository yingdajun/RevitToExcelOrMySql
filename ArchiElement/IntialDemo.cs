using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using Autodesk.Revit.UI;
namespace ArchiElement
{
  public  class IntialDemo
    {
        /// <summary>
        /// 导出EXCEL格式
        /// </summary>
        /// <param name="table"></param>
        /// <param name="file"></param>
        public static void dataTableToCsv(System.Data.DataTable table, string file)

        {

            string title = "";

            bool result=
            IsFileInUse(file);

            TaskDialog.Show("文件是否被占用",result.ToString());
            //这里报错
            FileStream fs = new FileStream(file, FileMode.OpenOrCreate);

            //FileStream fs1 = File.Open(file, FileMode.Open, FileAccess.Read);

            StreamWriter sw = new StreamWriter(new BufferedStream(fs), System.Text.Encoding.Default);

            for (int i = 0; i < table.Columns.Count; i++)

            {

                title += table.Columns[i].ColumnName + "\t"; //栏位：自动跳到下一单元格

            }

            title = title.Substring(0, title.Length - 1) + "\n";

            sw.Write(title);

            foreach (DataRow row in table.Rows)

            {

                string line = "";

                for (int i = 0; i < table.Columns.Count; i++)

                {

                    line += row[i].ToString().Trim() + "\t"; //内容：自动跳到下一单元格

                }

                line = line.Substring(0, line.Length - 1) + "\n";

                sw.Write(line);

            }

            sw.Close();

            fs.Close();

        }


        public static bool IsFileInUse(string fileName)
        {
            bool inUse = true;

            FileStream fs = null;
            try
            {

                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read,

                FileShare.None);

                inUse = false;
            }
            catch
            {

            }
            finally
            {
                if (fs != null)

                    fs.Close();
            }
            return inUse;//true表示正在使用,false没有使用
        }
    }
}

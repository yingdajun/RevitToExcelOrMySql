using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Diagnostics;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using MySql.Data;
using MySql.Data.MySqlClient;
namespace ArchiElement
{
    class MySQLIntial
    {
public static int BulkInsert(string connectionString, DataTable table)
        {
            if (string.IsNullOrEmpty(table.TableName))
                throw new Exception("请给DataTable的TableName属性附上表名称");
            if (table.Rows.Count == 0) return 0;
            int insertCount = 0;
            string tmpPath = Path.Combine(Directory.GetCurrentDirectory(), "Temp.csv");             string csv = DataTableToCsv(table);
            File.WriteAllText(tmpPath, csv);
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                try
                {
                    Stopwatch stopwatch = new Stopwatch();
                    stopwatch.Start();
                    conn.Open();
                                                            MySqlBulkLoader bulk = new MySqlBulkLoader(conn)
                    {
                        FieldTerminator = ",",
                        FieldQuotationCharacter = '"',
                        EscapeCharacter = '"',
                        LineTerminator = "\r\n",
                        FileName = tmpPath,
                        NumberOfLinesToSkip = 0,
                        TableName = table.TableName,
                    };
                    insertCount = bulk.Load();
                    stopwatch.Stop();
                                        TaskDialog.Show("元素属性批量导入的时间", "耗时:" + stopwatch.ElapsedMilliseconds.ToString());
                                    }
                catch (MySqlException ex)
                {
                                        MessageBox.Show("连接失败！", "测试结果");
                    TaskDialog.Show("REVIT", ex.ToString());
                    throw ex;
                }
            }
            File.Delete(tmpPath);
            return insertCount;
        }
public static int BulkInsert(string connectionString, DataTable table, int i)
        {
            if (string.IsNullOrEmpty(table.TableName))
                throw new Exception("请给DataTable的TableName属性附上表名称");
            if (table.Rows.Count == 0) return 0;
            int insertCount = 0;
            string tmpPath = Path.Combine(Directory.GetCurrentDirectory(), "Temp.csv");             string csv = DataTableToCsv(table);
            File.WriteAllText(tmpPath, csv);
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                try
                {
                    Stopwatch stopwatch = new Stopwatch();
                    stopwatch.Start();
                    conn.Open();
                    bool result =
                    AlterTableExample(conn, i);
                    MySqlBulkLoader bulk = new MySqlBulkLoader(conn)
                    {
                        FieldTerminator = ",",
                        FieldQuotationCharacter = '"',
                        EscapeCharacter = '"',
                        LineTerminator = "\r\n",
                        FileName = tmpPath,
                        NumberOfLinesToSkip = 0,
                        TableName = table.TableName,
                    };
                    insertCount = bulk.Load();
                    stopwatch.Stop();
                    conn.Close();
                    if (result)
                    {
                        TaskDialog.Show("批量导入数据库的时间", "耗时:" + stopwatch.ElapsedMilliseconds.ToString() + "毫秒");
                    }
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show("连接失败！", "测试结果");
                    TaskDialog.Show("REVIT", ex.ToString());
                    throw ex;
                }
            }
            File.Delete(tmpPath);
            return insertCount;
        }

 
 private static string DataTableToCsv(DataTable table)
        {
                                                StringBuilder sb = new StringBuilder();
            DataColumn colum;
            foreach (DataRow row in table.Rows)
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    colum = table.Columns[i];
                    if (i != 0) sb.Append(",");
                    if (colum.DataType == typeof(string) && row[colum].ToString().Contains(","))
                    {
                        sb.Append("\"" + row[colum].ToString().Replace("\"", "\"\"") + "\"");
                    }
                    else sb.Append(row[colum].ToString());
                }
                sb.AppendLine();
            }
            return sb.ToString();
        }
 public static bool AlterTableExample(MySqlConnection conn, int type)
        {
            bool result = false;
            string createStatement = null;
            string usedata = "use mytest";
            switch (type)
            {
                case 0:                     createStatement = "CREATE TABLE WallTable (id SMALLINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,walltype VarChar(50),levelName VarChar(50),width Decimal(10,2),height Decimal(10,2))";
                    break;
                case 1:                     createStatement = "CREATE TABLE FloorTable (id SMALLINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,floorType VarChar(50),levelName VarChar(50),thick Decimal(10,2) ,area Decimal(10,2),offset Decimal(10,2))";
                    break;
                case 2:                     createStatement = "CREATE TABLE DoorTable (id SMALLINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,doortype VarChar(50),levelName VarChar(50),drma VarChar(50),kjma VarChar(50),width VarChar(50),bottomheight VarChar(50),height VarChar(50))";
                    break;
                case 3:                     createStatement = "CREATE TABLE WindowTable (id SMALLINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,windowsType VarChar(50),levelName VarChar(50),outcha VarChar(50),incha VarChar(50),bolicha VarChar(50),width VarChar(50),bottomheight VarChar(50),height VarChar(50))";
                    break;
                case 4:                     createStatement = "CREATE TABLE RoomTable (id SMALLINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,number VarChar(50),name VarChar(50),levelName VarChar(50),bottomOffset Decimal(10,2),area Decimal(10,2),topoffset Decimal(10,2))";
                    break;
                case 5:                     createStatement = "CREATE TABLE StructColumnTable (id SMALLINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,columnType VarChar(50),category VarChar(50),type VarChar(50),material VarChar(50),width VarChar(50), height  VarChar(50), bottomLev VarChar(50), bottomOffset  VarChar(50), topLev VarChar(50), topOffset  VarChar(50),isMove VARCHAR(50))";
                    break;
                case 6:                    createStatement = "CREATE TABLE BeamTable (id SMALLINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,beamName VarChar(50),Olev VarChar(50),material VarChar(50),length Decimal(10,2),y_duizheng VarChar(50),y_offset Decimal(10,2),yz_duizheng VarChar(50),z_duizheng VarChar(50),z_offset Decimal(10,2))";
                    break;
                case 7:                     createStatement = "CREATE TABLE DuctTable (id SMALLINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,DuctName VarChar(50),Olev VarChar(50),Width  Decimal(10,2),Height Decimal(10,2),Offset Decimal(10,2),shuidui VarChar(50),chuizhidui VarChar(50),type VarChar(50))";
                    break;
                case 8:                     createStatement = "CREATE TABLE PipeTable (id SMALLINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,pipeName VarChar(50),Olev VarChar(50),shuidui VarChar(50),chuizhidui VarChar(50),offset Decimal(10,2),slope VarChar(50),type VarChar(50),guanduan VarChar(50),r VarChar(50))";
                    break;
                case 9:                    createStatement = "CREATE TABLE EletricalQTable (id SMALLINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,categoryName VarChar(50),Olev VarChar(50),shuidui VarChar(50),chuizhidui VarChar(50),offset Decimal(10,2),width VarChar(50),height VarChar(50),hendang VarChar(50),familyName VarChar(50))";
                    break;
            }
                        try
            {
                                using (MySqlCommand cmdb = new MySqlCommand(usedata, conn))
                {
                    cmdb.ExecuteNonQuery();
                }
                                using (MySqlCommand cmd = new MySqlCommand(createStatement, conn))
                {
                                        cmd.ExecuteNonQuery();
                }
                TaskDialog.Show("建表成功", "建表成功");
                result = true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("建表失败", "建表失败");
                TaskDialog.Show("建表失败的原因", ex.Message);
            }
            return result;
        }
    }
}

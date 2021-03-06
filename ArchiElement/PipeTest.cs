﻿using System;
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
namespace ArchiElement
{
    [Transaction(TransactionMode.Manual)]
    class PipeTest : IExternalCommand
    {
        public string excelPath = @"D:\管道.xls";
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
                        UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            FilteredElementCollector pipeCollector = PipeIntial.PipeCollector(doc);
            DataTable dt = null;
               dt= PipeIntial.CreatePipeExcelTitle();
                        PipeIntial.
            PipeElementExcelPara(doc, pipeCollector, dt);
            TaskDialog.Show("EXCEL放置位置", excelPath.ToString());
            LChuDemo.
            dataTableToCsv(dt, excelPath);
            System.Diagnostics.Process.Start(excelPath);
                        dt = null;
                        dt = PipeIntial.CreatePipeMySQLTitle();
                        dt.TableName = "PipeTable";
                        PipeIntial.PipeElementMYSQLPara(doc, pipeCollector, dt);
            string connStr = "server=localhost;database=mytest;uid=root;pwd=123456";
            var result = MySQLIntial.BulkInsert(connStr, dt,8);
                        if (result != 0.0)
            {
                TaskDialog.Show("导出到MYSQL中成功", "数据已经存入" + "数据库mytest" + dt.TableName + "中");
            }
            return Result.Succeeded;
        }
                                                                                                                                                                                                                                                            }
}

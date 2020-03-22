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
namespace ArchiElement
{
    [Transaction(TransactionMode.Manual)]
    class BeamTest : IExternalCommand
    {
        string excelPath = @"D:\梁.xls";
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            FilteredElementCollector beamCollector = BeamIntial.BeamCollector(doc);
            DataTable dt = BeamIntial.CreateBeamExcelTitle();
            BeamIntial.
            BeamElementExcelPara(doc, beamCollector, dt);
            LChuDemo.dataTableToCsv(dt, excelPath);
            string savePath =
            LChuDemo.PickFolderInfo(excelPath);
            System.Diagnostics.Process.Start(savePath);
            dt = BeamIntial.CreateBeamMySQLTitle();
            dt.TableName = "BeamTable";
            BeamIntial.BeamElementMySQLPara(doc, beamCollector, dt);
            string connStr = "server=localhost;database=mytest;uid=root;pwd=123456";
            var result = MySQLIntial.BulkInsert(connStr, dt, 6);
            if (result != 0.0)
            {
                TaskDialog.Show("导出到MYSQL中成功", "数据已经存入" + "数据库mytest" + "/n" + dt.TableName + "中");
            }
            return Result.Succeeded;
        }
    }
}

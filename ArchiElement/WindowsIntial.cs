using Autodesk.Revit.DB;
using System;
using System.Data;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using System.Linq;
namespace ArchiElement
{
    class WindowsIntial
    {
        public static FilteredElementCollector WindowsCollector(Document doc)
        {
            FilteredElementCollector doorCollector = new FilteredElementCollector(doc);
            doorCollector.OfCategory(BuiltInCategory.OST_Windows).OfClass(typeof(FamilyInstance));
            TaskDialog.Show("REVIT", doorCollector.Count().ToString());
            return doorCollector;
        }
        public static void WindowsElementExcelPara(Document doc, FilteredElementCollector doorCollector,
System.Data.DataTable dt)
        {
            foreach (Element ele in doorCollector)
            {
                FamilyInstance windows = ele as FamilyInstance;
                Level level = level = doc.GetElement(ele.LevelId) as Level;
                FamilySymbol familySymbol =
    windows.Document.GetElement(windows.GetTypeId()) as FamilySymbol;
                if (windows != null)
                {
                    string bottomheight = null;
                    string width = null;
                    string height = null;
                    string bolicha = null;
                    string outcha = null;
                    string intercha = null;
                    foreach (Parameter param in windows.Parameters)
                    {
                        InternalDefinition definition = null;
                        definition = param.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                        if (BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM == definition.BuiltInParameter)
                        {
                            bottomheight = param.AsValueString();
                        }
                    }
                    foreach (Parameter para in familySymbol.ParametersMap)
                    {
                        InternalDefinition definition = para.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                        if (BuiltInParameter.DOOR_WIDTH == definition.BuiltInParameter)
                        {
                            width = para.AsValueString();
                        }
                        if (BuiltInParameter.GENERIC_HEIGHT == definition.BuiltInParameter)
                        {
                            height = para.AsValueString();
                        }
                        if (definition.Name == "玻璃嵌板材质")
                        {
                            bolicha = para.AsValueString();
                        }
                        if (definition.Name == "框架内部材质")
                        {
                            intercha = para.AsValueString();
                        }
                        if (definition.Name == "框架内部材质")
                        {
                            outcha = para.AsValueString();
                        }
                    }
                    CreateWindowExcelRow(dt, windows.Symbol.FamilyName, level.Name, outcha, intercha,
                        bolicha, width, bottomheight, height);
                }
            }
        }
        public static void WindowsElementMySQLPara(Document doc, FilteredElementCollector doorCollector,
            System.Data.DataTable dt)
        {
            foreach (Element ele in doorCollector)
            {
                FamilyInstance windows = ele as FamilyInstance;
                Level level = level = doc.GetElement(ele.LevelId) as Level;
                FamilySymbol familySymbol =
    windows.Document.GetElement(windows.GetTypeId()) as FamilySymbol;
                if (windows != null)
                {
                    string bottomheight = null;
                    string width = null;
                    string height = null;
                    string bolicha = null;
                    string outcha = null;
                    string intercha = null;
                    foreach (Parameter param in windows.Parameters)
                    {
                        InternalDefinition definition = null;
                        definition = param.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                        if (BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM == definition.BuiltInParameter)
                        {
                            bottomheight = param.AsValueString();
                        }
                    }
                    foreach (Parameter para in familySymbol.ParametersMap)
                    {
                        InternalDefinition definition = para.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                        if (BuiltInParameter.DOOR_WIDTH == definition.BuiltInParameter)
                        {
                            width = para.AsValueString();
                        }
                        if (BuiltInParameter.GENERIC_HEIGHT == definition.BuiltInParameter)
                        {
                            height = para.AsValueString();
                        }
                        if (definition.Name == "玻璃嵌板材质")
                        {
                            bolicha = para.AsValueString();
                        }
                        if (definition.Name == "框架内部材质")
                        {
                            intercha = para.AsValueString();
                        }
                        if (definition.Name == "框架内部材质")
                        {
                            outcha = para.AsValueString();
                        }
                    }
                    CreateWindowMySQLRow(dt, windows.Symbol.FamilyName, level.Name, outcha, intercha,
                        bolicha, width, bottomheight, height);
                }
            }
        }
        public static void CreateWindowExcelRow(System.Data.DataTable dt, string windowsType, string levelName,
string outcha, string incha, string bolicha, string width, string bottomheight, string height)
        {
            DataRow dr = dt.NewRow();
            dr["窗类型"] = windowsType;
            dr["标高"] = levelName;
            dr["框架外部材质"] = outcha;
            dr["框架内部材质"] = incha;
            dr["玻璃嵌板材质"] = incha;
            dr["宽度"] = width;
            dr["底高度"] = bottomheight;
            dr["高度"] = height;
            dt.Rows.Add(dr);
            dr = null;
        }
        public static System.Data.DataTable CreateWindowsExcelTitle()
        {
            System.Data.DataTable dt = new System.Data.DataTable("窗信息表");
            DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true; dc.AutoIncrementSeed = 1; dc.AutoIncrementStep = 1; dc.AllowDBNull = false; dt.Columns.Add(new DataColumn("窗类型", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("标高", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("框架外部材质", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("框架内部材质", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("玻璃嵌板材质", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("宽度", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("底高度", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("高度", Type.GetType("System.String")));
            return dt;
        }
        public static void CreateWindowMySQLRow(System.Data.DataTable dt, string windowsType, string levelName,
string outcha, string incha, string bolicha, string width, string bottomheight, string height)
        {
            DataRow dr = dt.NewRow();
            dr["windowsType"] = windowsType;
            dr["levelName"] = levelName;
            dr["outcha"] = outcha;
            dr["incha"] = incha;
            dr["incha"] = incha;
            dr["width"] = width;
            dr["bottomheight"] = bottomheight;
            dr["height"] = height;
            dt.Rows.Add(dr);
            dr = null;
        }
        public static System.Data.DataTable CreateWindowsMySQLTitle()
        {
            System.Data.DataTable dt = new System.Data.DataTable("窗信息表");
            DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true; dc.AutoIncrementSeed = 1; dc.AutoIncrementStep = 1; dc.AllowDBNull = false; dt.Columns.Add(new DataColumn("windowsType", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("levelName", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("outcha", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("incha", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("bolicha", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("width", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("bottomheight", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("height", Type.GetType("System.String")));
            return dt;
        }
        public static double FeetTomm(double val)
        {
            return val * 304.8;
        }
    }
}

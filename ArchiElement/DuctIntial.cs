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
using Autodesk.Revit.DB.Mechanical;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Configuration;
using System.Diagnostics;
using System.IO;
namespace ArchiElement
{
    class DuctIntial
    {
        public static FilteredElementCollector DuctCollector(Document doc)
        {
            FilteredElementCollector ductCollector = new FilteredElementCollector(doc);
            ductCollector.OfCategory(BuiltInCategory.OST_DuctCurves).OfClass(typeof(Duct));
            TaskDialog.Show("REVIT", ductCollector.Count().ToString());
            return ductCollector;
        }
        public static System.Data.DataTable CreateDuctExcelTitle()
        {
            System.Data.DataTable dt = new System.Data.DataTable("风管信息表");
            DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true; dc.AutoIncrementSeed = 1; dc.AutoIncrementStep = 1; dc.AllowDBNull = false; dt.Columns.Add(new DataColumn("风管名称", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("参照标高", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("宽度", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("高度", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("偏移", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("水平对正", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("垂直对正", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("系统类型", Type.GetType("System.String")));
            return dt;
        }
        public static System.Data.DataTable CreateDuctMySQLTitle()
        {
            System.Data.DataTable dt = new System.Data.DataTable("风管信息表");
            DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true; dc.AutoIncrementSeed = 1; dc.AutoIncrementStep = 1; dc.AllowDBNull = false; dt.Columns.Add(new DataColumn("DuctName", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Olev", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Width", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("Height", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("Offset", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("shuidui", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("chuizhidui", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("type", Type.GetType("System.String")));
            return dt;
        }
        public static void DuctElementExcelPara(Document doc, FilteredElementCollector ductCollector,
System.Data.DataTable dt)
        {
            foreach (Element ele in ductCollector)
            {
                Duct duct = ele as Duct;
                if (duct != null)
                {
                    string Olev = null;
                    double width = 0.0;
                    double height = 0.0;
                    double offset = 0.0;
                    string shuidui = null;
                    string chuizhidui = null;
                    string type = null;
                    foreach (Parameter param in duct.Parameters)
                    {
                        InternalDefinition definition = null;
                        definition = param.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                        if (BuiltInParameter.RBS_START_LEVEL_PARAM == definition.BuiltInParameter)
                        {
                            Olev = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_CURVE_WIDTH_PARAM == definition.BuiltInParameter)
                        {
                            width = param.AsDouble();
                        }
                        if (BuiltInParameter.RBS_CURVE_HEIGHT_PARAM == definition.BuiltInParameter)
                        {
                            height = param.AsDouble();
                        }
                        if (BuiltInParameter.RBS_OFFSET_PARAM == definition.BuiltInParameter)
                        {
                            offset = param.AsDouble();
                        }
                        if (BuiltInParameter.RBS_CURVE_HOR_OFFSET_PARAM == definition.BuiltInParameter)
                        {
                            shuidui = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_CURVE_VERT_OFFSET_PARAM == definition.BuiltInParameter)
                        {
                            chuizhidui = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM == definition.BuiltInParameter)
                        {
                            type = param.AsValueString();
                        }
                    }
                    CreateDuctEXCELRow(dt, duct.DuctType.Name, Olev,
                        FeetTomm(width), FeetTomm(height), FeetTomm(offset),
                        shuidui, chuizhidui, type);
                }
            }
        }
        public static void DuctElementMYSQLPara(Document doc, FilteredElementCollector ductCollector,
        System.Data.DataTable dt)
        {
            foreach (Element ele in ductCollector)
            {
                Duct duct = ele as Duct;
                if (duct != null)
                {
                    string Olev = null;
                    double width = 0.0;
                    double height = 0.0;
                    double offset = 0.0;
                    string shuidui = null;
                    string chuizhidui = null;
                    string type = null;
                    foreach (Parameter param in duct.Parameters)
                    {
                        InternalDefinition definition = null;
                        definition = param.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                        if (BuiltInParameter.RBS_START_LEVEL_PARAM == definition.BuiltInParameter)
                        {
                            Olev = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_CURVE_WIDTH_PARAM == definition.BuiltInParameter)
                        {
                            width = param.AsDouble();
                        }
                        if (BuiltInParameter.RBS_CURVE_HEIGHT_PARAM == definition.BuiltInParameter)
                        {
                            height = param.AsDouble();
                        }
                        if (BuiltInParameter.RBS_OFFSET_PARAM == definition.BuiltInParameter)
                        {
                            offset = param.AsDouble();
                        }
                        if (BuiltInParameter.RBS_CURVE_HOR_OFFSET_PARAM == definition.BuiltInParameter)
                        {
                            shuidui = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_CURVE_VERT_OFFSET_PARAM == definition.BuiltInParameter)
                        {
                            chuizhidui = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM == definition.BuiltInParameter)
                        {
                            type = param.AsValueString();
                        }
                    }
                    CreateDuctMYSQLRow(dt, duct.DuctType.Name, Olev,
                        FeetTomm(width), FeetTomm(height), FeetTomm(offset),
                        shuidui, chuizhidui, type);
                }
            }
        }
        private static void CreateDuctEXCELRow(DataTable dt, string DuctName, string Olev,
           double width, double height, double offset, string shuidui, string chuizhidui, string type)
        {
            DataRow dr = dt.NewRow();
            dr["风管名称"] = DuctName;
            dr["参照标高"] = Olev;
            dr["宽度"] = width;
            dr["高度"] = height;
            dr["偏移"] = offset;
            dr["水平对正"] = shuidui;
            dr["垂直对正"] = chuizhidui;
            dr["系统类型"] = type;
            dt.Rows.Add(dr);
            dr = null;
        }
        private static void CreateDuctMYSQLRow(DataTable dt, string DuctName, string Olev,
           double width, double height, double offset, string shuidui, string chuizhidui, string type)
        {
            DataRow dr = dt.NewRow();
            dr["DuctName"] = DuctName;
            dr["Olev"] = Olev;
            dr["Width"] = width;
            dr["Height"] = height;
            dr["Offset"] = offset;
            dr["shuidui"] = shuidui;
            dr["chuizhidui"] = chuizhidui;
            dr["type"] = type;
            dt.Rows.Add(dr);
            dr = null;
        }
        public static double FeetTomm(double val)
        {
            return val * 304.8;
        }
    }
}

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
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
namespace ArchiElement
{
    class EletricalQIntial
    {
        public static FilteredElementCollector EletricalQCollector(Document doc)
        {
            FilteredElementCollector qiaojiaCollector = new FilteredElementCollector(doc);
            qiaojiaCollector.OfCategory(BuiltInCategory.OST_CableTray).OfClass(typeof(CableTray));
            TaskDialog.Show("REVIT", qiaojiaCollector.Count().ToString());
            return qiaojiaCollector;
        }
        public static void EletricalQElementExcelPara(Document doc, FilteredElementCollector cableCollector,
System.Data.DataTable dt)
        {
            foreach (Element ele in cableCollector)
            {
                CableTray cable = ele as CableTray;
                Level level = level = doc.GetElement(cable.LevelId) as Level;
                if (cable != null)
                {
                    string Olev = null;
                    string shuidui = null;
                    string chuizhidui = null;
                    double offset = 0.0;
                    string width = null;
                    string height = null;
                    string hendang = null;
                    string familyname = null;
                    foreach (Parameter param in cable.Parameters)
                    {
                        InternalDefinition definition = null;
                        definition = param.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                        if (BuiltInParameter.RBS_START_LEVEL_PARAM == definition.BuiltInParameter)
                        {
                            Olev = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_CURVE_HOR_OFFSET_PARAM == definition.BuiltInParameter)
                        {
                            shuidui = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_CURVE_VERT_OFFSET_PARAM == definition.BuiltInParameter)
                        {
                            chuizhidui = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_OFFSET_PARAM == definition.BuiltInParameter)
                        {
                            offset = FeetTomm(param.AsDouble());
                        }
                        if (BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM == definition.BuiltInParameter)
                        {
                            width = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM == definition.BuiltInParameter)
                        {
                            height = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_CABLETRAY_RUNGSPACE == definition.BuiltInParameter)
                        {
                            hendang = param.AsValueString();
                        }
                        if (BuiltInParameter.ELEM_FAMILY_PARAM == definition.BuiltInParameter)
                        {
                            familyname = param.AsValueString();
                        }
                    }
                    CreateEletricalQExcelRow(dt, cable.Category.Name, Olev,
                        shuidui, chuizhidui,
                        offset, width, height, hendang, familyname
                        );
                }
            }
        }
        public static void EletricalQElementMySQLPara(Document doc, FilteredElementCollector cableCollector,
               System.Data.DataTable dt)
        {
            foreach (Element ele in cableCollector)
            {
                CableTray cable = ele as CableTray;
                Level level = level = doc.GetElement(cable.LevelId) as Level;
                if (cable != null)
                {
                    string Olev = null;
                    string shuidui = null;
                    string chuizhidui = null;
                    double offset = 0.0;
                    string width = null;
                    string height = null;
                    string hendang = null;
                    string familyname = null;
                    foreach (Parameter param in cable.Parameters)
                    {
                        InternalDefinition definition = null;
                        definition = param.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                        if (BuiltInParameter.RBS_START_LEVEL_PARAM == definition.BuiltInParameter)
                        {
                            Olev = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_CURVE_HOR_OFFSET_PARAM == definition.BuiltInParameter)
                        {
                            shuidui = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_CURVE_VERT_OFFSET_PARAM == definition.BuiltInParameter)
                        {
                            chuizhidui = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_OFFSET_PARAM == definition.BuiltInParameter)
                        {
                            offset = FeetTomm(param.AsDouble());
                        }
                        if (BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM == definition.BuiltInParameter)
                        {
                            width = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM == definition.BuiltInParameter)
                        {
                            height = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_CABLETRAY_RUNGSPACE == definition.BuiltInParameter)
                        {
                            hendang = param.AsValueString();
                        }
                        if (BuiltInParameter.ELEM_FAMILY_PARAM == definition.BuiltInParameter)
                        {
                            familyname = param.AsValueString();
                        }
                    }
                    CreateEletricalQMySQLRow(dt, cable.Category.Name, Olev,
                        shuidui, chuizhidui,
                        offset, width, height, hendang, familyname
                        );
                }
            }
        }
        private static void CreateEletricalQExcelRow(DataTable dt, string categoryName, string Olev,
            string shuidui, string chuizhidui, double offset, string width, string height, string hendang
, string familyname)
        {
            DataRow dr = dt.NewRow();
            dr["桥架名称"] = categoryName;
            dr["参照标高"] = Olev;
            dr["水平对正"] = shuidui;
            dr["垂直对正"] = chuizhidui;
            dr["偏移"] = offset;
            dr["宽度"] = width;
            dr["高度"] = height;
            dr["横档间距"] = hendang;
            dr["族名称"] = familyname;
            dt.Rows.Add(dr);
            dr = null;
        }
        public static System.Data.DataTable CreateEletricalQExcelTitle()
        {
            System.Data.DataTable dt = new System.Data.DataTable("桥架信息表");
            DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true; dc.AutoIncrementSeed = 1; dc.AutoIncrementStep = 1; dc.AllowDBNull = false; dt.Columns.Add(new DataColumn("桥架名称", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("参照标高", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("水平对正", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("垂直对正", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("偏移", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("宽度", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("高度", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("横档间距", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("族名称", Type.GetType("System.String")));
            return dt;
        }
        private static void CreateEletricalQMySQLRow(DataTable dt, string categoryName, string Olev,
          string shuidui, string chuizhidui, double offset, string width, string height, string hendang
, string familyName)
        {
            DataRow dr = dt.NewRow();
            dr["categoryName"] = categoryName;
            dr["Olev"] = Olev;
            dr["shuidui"] = shuidui;
            dr["chuizhidui"] = chuizhidui;
            dr["offset"] = offset;
            dr["width"] = width;
            dr["height"] = height;
            dr["hendang"] = hendang;
            dr["familyName"] = familyName;
            dt.Rows.Add(dr);
            dr = null;
        }
        public static System.Data.DataTable CreateEletricalQMySQLTitle()
        {
            System.Data.DataTable dt = new System.Data.DataTable("桥架信息表");
            DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true; dc.AutoIncrementSeed = 1; dc.AutoIncrementStep = 1; dc.AllowDBNull = false; dt.Columns.Add(new DataColumn("categoryName", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Olev", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("shuidui", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("chuizhidui", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("offset", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("width", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("height", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("hendang", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("familyName", Type.GetType("System.String")));
            return dt;
        }
        public static double FeetTomm(double val)
        {
            return val * 304.8;
        }
    }
}

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
    class ColumnIntial
    {
        public static FilteredElementCollector ColumnCollector(Document doc)
        {
            FilteredElementCollector columnCollector = new FilteredElementCollector(doc);
            columnCollector.OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilyInstance));
            TaskDialog.Show("REVIT", columnCollector.Count().ToString());
            return columnCollector;
        }
        public static void ColumnElementExcelPara(Document doc, FilteredElementCollector doorCollector,
System.Data.DataTable dt)
        {
            foreach (Element ele in doorCollector)
            {
                FamilyInstance column = ele as FamilyInstance;
                FamilySymbol familySymbol =
column.Document.GetElement(column.GetTypeId()) as FamilySymbol;
                if (column != null)
                {
                    string width = null;
                    string height = null;
                    string bottomLev = null;
                    string bottomOffset = null;
                    string topLev = null;
                    string topOffset = null;
                    string matrial = null;
                    string type = null;
                    string isMove = null;
                    string category = null;
                    foreach (Parameter param in column.Parameters)
                    {
                        InternalDefinition definition = null;
                        definition = param.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                        if (BuiltInParameter.ELEM_CATEGORY_PARAM_MT == definition.BuiltInParameter)
                        {
                            category = param.AsValueString();
                        }
                        if (BuiltInParameter.FAMILY_BASE_LEVEL_PARAM == definition.BuiltInParameter)
                        {
                            bottomLev = param.AsValueString();
                        }
                        if (BuiltInParameter.SCHEDULE_BASE_LEVEL_OFFSET_PARAM == definition.BuiltInParameter)
                        {
                            bottomOffset = param.AsValueString();
                        }
                        if (BuiltInParameter.FAMILY_TOP_LEVEL_PARAM == definition.BuiltInParameter)
                        {
                            topLev = param.AsValueString();
                        }
                        if (BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM == definition.BuiltInParameter)
                        {
                            topOffset = param.AsValueString();
                        }
                        if (BuiltInParameter.STRUCTURAL_MATERIAL_PARAM == definition.BuiltInParameter)
                        {
                            matrial = param.AsValueString();
                        }
                        if (BuiltInParameter.SLANTED_COLUMN_TYPE_PARAM == definition.BuiltInParameter)
                        {
                            type = param.AsValueString();
                        }
                        if (BuiltInParameter.INSTANCE_MOVES_WITH_GRID_PARAM == definition.BuiltInParameter)
                        {
                            isMove = param.AsValueString();
                        }
                    }
                    foreach (Parameter para in familySymbol.ParametersMap)
                    {
                        InternalDefinition definition = para.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                        if (BuiltInParameter.STRUCTURAL_SECTION_COMMON_WIDTH == definition.BuiltInParameter)
                        {
                            width = para.AsValueString();
                        }
                        if (BuiltInParameter.STRUCTURAL_SECTION_COMMON_HEIGHT == definition.BuiltInParameter)
                        {
                            height = para.AsValueString();
                        }
                    }
                    CreateColumnExcelRow(dt, column.Symbol.FamilyName,
                        bottomLev, bottomOffset, topLev, topOffset, matrial, type, isMove,
                        category, width, height);
                }
            }
        }
        public static void ColumnElementMySQLPara(Document doc, FilteredElementCollector doorCollector,
           System.Data.DataTable dt)
        {
            foreach (Element ele in doorCollector)
            {
                FamilyInstance column = ele as FamilyInstance;
                FamilySymbol familySymbol =
column.Document.GetElement(column.GetTypeId()) as FamilySymbol;
                if (column != null)
                {
                    string width = null;
                    string height = null;
                    string bottomLev = null;
                    string bottomOffset = null;
                    string topLev = null;
                    string topOffset = null;
                    string matrial = null;
                    string type = null;
                    string isMove = null;
                    string category = null;
                    foreach (Parameter param in column.Parameters)
                    {
                        InternalDefinition definition = null;
                        definition = param.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                        if (BuiltInParameter.ELEM_CATEGORY_PARAM_MT == definition.BuiltInParameter)
                        {
                            category = param.AsValueString();
                        }
                        if (BuiltInParameter.FAMILY_BASE_LEVEL_PARAM == definition.BuiltInParameter)
                        {
                            bottomLev = param.AsValueString();
                        }
                        if (BuiltInParameter.SCHEDULE_BASE_LEVEL_OFFSET_PARAM == definition.BuiltInParameter)
                        {
                            bottomOffset = param.AsValueString();
                        }
                        if (BuiltInParameter.FAMILY_TOP_LEVEL_PARAM == definition.BuiltInParameter)
                        {
                            topLev = param.AsValueString();
                        }
                        if (BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM == definition.BuiltInParameter)
                        {
                            topOffset = param.AsValueString();
                        }
                        if (BuiltInParameter.STRUCTURAL_MATERIAL_PARAM == definition.BuiltInParameter)
                        {
                            matrial = param.AsValueString();
                        }
                        if (BuiltInParameter.SLANTED_COLUMN_TYPE_PARAM == definition.BuiltInParameter)
                        {
                            type = param.AsValueString();
                        }
                        if (BuiltInParameter.INSTANCE_MOVES_WITH_GRID_PARAM == definition.BuiltInParameter)
                        {
                            isMove = param.AsValueString();
                        }
                    }
                    foreach (Parameter para in familySymbol.ParametersMap)
                    {
                        InternalDefinition definition = para.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                        if (BuiltInParameter.STRUCTURAL_SECTION_COMMON_WIDTH == definition.BuiltInParameter)
                        {
                            width = para.AsValueString();
                        }
                        if (BuiltInParameter.STRUCTURAL_SECTION_COMMON_HEIGHT == definition.BuiltInParameter)
                        {
                            height = para.AsValueString();
                        }
                    }
                    CreateColumnMySQLRow(dt, column.Symbol.FamilyName,
                        bottomLev, bottomOffset, topLev, topOffset, matrial, type, isMove,
                        category, width, height);
                }
            }
        }
        public static void CreateColumnExcelRow(System.Data.DataTable dt, string columnType,
string bottomLev, string bottomOffset, string topLev, string topOffset,
string matrial, string type, string isMove,
string category, string width, string height
)
        {
            DataRow dr = dt.NewRow();
            dr["柱类型"] = columnType;
            dr["柱类别"] = category;
            dr["柱样式"] = type;
            dr["结构材质"] = matrial;
            dr["随轴网移动"] = isMove;
            dr["底部标高"] = bottomLev;
            dr["底部偏移"] = bottomOffset;
            dr["顶部标高"] = topLev;
            dr["顶部偏移"] = topOffset;
            dr["宽度"] = width;
            dr["高度"] = height;
            dt.Rows.Add(dr);
            dr = null;
        }
        public static System.Data.DataTable CreateColumnExcelTitle()
        {
            System.Data.DataTable dt = new System.Data.DataTable("柱信息表");
            DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true; dc.AutoIncrementSeed = 1; dc.AutoIncrementStep = 1; dc.AllowDBNull = false; dt.Columns.Add(new DataColumn("柱类别", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("柱类型", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("柱样式", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("结构材质", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("宽度", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("高度", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("底部标高", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("底部偏移", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("顶部标高", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("顶部偏移", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("随轴网移动", Type.GetType("System.String")));
            return dt;
        }
        public static void CreateColumnMySQLRow(System.Data.DataTable dt, string columnType,
string bottomLev, string bottomOffset, string topLev, string topOffset,
string matrial, string type, string isMove,
string category, string width, string height
)
        {
            DataRow dr = dt.NewRow();
            dr["columnType"] = columnType;
            dr["category"] = category;
            dr["type"] = type;
            dr["matrial"] = matrial;
            dr["isMove"] = isMove;
            dr["bottomLev"] = bottomLev;
            dr["bottomOffset"] = bottomOffset;
            dr["topLev"] = topLev;
            dr["topOffset"] = topOffset;
            dr["width"] = width;
            dr["height"] = height;
            dt.Rows.Add(dr);
            dr = null;
        }
        public static System.Data.DataTable CreateColumnMySQLTitle()
        {
            System.Data.DataTable dt = new System.Data.DataTable("柱信息表");
            DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true; dc.AutoIncrementSeed = 1; dc.AutoIncrementStep = 1; dc.AllowDBNull = false; dt.Columns.Add(new DataColumn("columnType", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("category", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("type", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("matrial", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("width", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("height", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("bottomLev", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("bottomOffset", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("topLev", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("topOffset", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("isMove", Type.GetType("System.String")));
            return dt;
        }
        public static double FeetTomm(double val)
        {
            return val * 304.8;
        }
    }
}

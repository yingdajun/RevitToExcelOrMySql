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
namespace ArchiElement
{
    class PipeIntial
    {
        public static FilteredElementCollector PipeCollector(Document doc)
        {
            FilteredElementCollector pipeCollector = new FilteredElementCollector(doc);
            pipeCollector.OfCategory(BuiltInCategory.OST_PipeCurves).OfClass(typeof(Pipe));
            TaskDialog.Show("REVIT", pipeCollector.Count().ToString());
            return pipeCollector;
        }
        public static void PipeElementExcelPara(Document doc, FilteredElementCollector ductCollector,
System.Data.DataTable dt)
        {
            foreach (Element ele in ductCollector)
            {
                Pipe pipe = ele as Pipe;
                Level level = level = doc.GetElement(pipe.LevelId) as Level;
                if (pipe != null)
                {
                    string Olev = null;
                    string shuidui = null;
                    string chuizhidui = null;
                    double offset = 0.0;
                    string slope = null;
                    string type = null;
                    string guanduan = null;
                    string r = null;
                    foreach (Parameter param in pipe.Parameters)
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
                            offset = param.AsDouble();
                        }
                        if (BuiltInParameter.RBS_PIPE_SLOPE == definition.BuiltInParameter)
                        {
                            slope = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM == definition.BuiltInParameter)
                        {
                            type = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_PIPE_SEGMENT_PARAM == definition.BuiltInParameter)
                        {
                            guanduan = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_PIPE_DIAMETER_PARAM == definition.BuiltInParameter)
                        {
                            r = param.AsValueString();
                        }
                    }
                    CreatePipeExcelRow(dt, pipe.PipeType.Name, Olev,
                        shuidui, chuizhidui, FeetTomm(offset),
                        slope, type, guanduan, r
                        );
                }
            }
        }
        public static void PipeElementMYSQLPara(Document doc, FilteredElementCollector ductCollector,
        System.Data.DataTable dt)
        {
            foreach (Element ele in ductCollector)
            {
                Pipe pipe = ele as Pipe;
                Level level = level = doc.GetElement(pipe.LevelId) as Level;
                if (pipe != null)
                {
                    string Olev = null;
                    string shuidui = null;
                    string chuizhidui = null;
                    double offset = 0.0;
                    string slope = null;
                    string type = null;
                    string guanduan = null;
                    string r = null;
                    foreach (Parameter param in pipe.Parameters)
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
                            offset = param.AsDouble();
                        }
                        if (BuiltInParameter.RBS_PIPE_SLOPE == definition.BuiltInParameter)
                        {
                            slope = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM == definition.BuiltInParameter)
                        {
                            type = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_PIPE_SEGMENT_PARAM == definition.BuiltInParameter)
                        {
                            guanduan = param.AsValueString();
                        }
                        if (BuiltInParameter.RBS_PIPE_DIAMETER_PARAM == definition.BuiltInParameter)
                        {
                            r = param.AsValueString();
                        }
                    }
                    CreatePipeMySQLRow(dt, pipe.PipeType.Name, Olev,
    shuidui, chuizhidui, FeetTomm(offset),
    slope, type, guanduan, r
    );
                }
            }
        }
        private static void CreatePipeExcelRow(DataTable dt, string familyName, string Olev,
string shuidui, string chuizhidui, double offset, string slope, string type, string guanduan,
string r
)
        {
            DataRow dr = dt.NewRow();
            dr["管道名称"] = familyName;
            dr["参照标高"] = Olev;
            dr["水平对正"] = shuidui;
            dr["垂直对正"] = chuizhidui;
            dr["偏移"] = offset;
            dr["坡度"] = slope;
            dr["系统类型"] = type;
            dr["管段"] = guanduan;
            dr["直径"] = r;
            dt.Rows.Add(dr);
            dr = null;
        }
        public static System.Data.DataTable CreatePipeExcelTitle()
        {
            System.Data.DataTable dt = new System.Data.DataTable("管道信息表");
            DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true; dc.AutoIncrementSeed = 1; dc.AutoIncrementStep = 1; dc.AllowDBNull = false; dt.Columns.Add(new DataColumn("管道名称", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("参照标高", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("水平对正", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("垂直对正", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("偏移", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("坡度", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("系统类型", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("管段", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("直径", Type.GetType("System.String")));
            return dt;
        }
        private static void CreatePipeMySQLRow(DataTable dt, string pipeName, string Olev,
string shuidui, string chuizhidui, double offset, string slope, string type, string guanduan,
string r
)
        {
            DataRow dr = dt.NewRow();
            dr["pipeName"] = pipeName;
            dr["Olev"] = Olev;
            dr["shuidui"] = shuidui;
            dr["chuizhidui"] = chuizhidui;
            dr["offset"] = offset;
            dr["slope"] = slope;
            dr["type"] = type;
            dr["guanduan"] = guanduan;
            dr["r"] = r;
            dt.Rows.Add(dr);
            dr = null;
        }
        public static System.Data.DataTable CreatePipeMySQLTitle()
        {
            System.Data.DataTable dt = new System.Data.DataTable("管道信息表");
            DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true; dc.AutoIncrementSeed = 1; dc.AutoIncrementStep = 1; dc.AllowDBNull = false; dt.Columns.Add(new DataColumn("pipeName", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Olev", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("shuidui", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("chuizhidui", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("offset", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("slope", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("type", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("guanduan", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("r", Type.GetType("System.String")));
            return dt;
        }
        public static double FeetTomm(double val)
        {
            return val * 304.8;
        }
    }
}

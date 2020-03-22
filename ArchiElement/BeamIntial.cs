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
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
namespace ArchiElement
{
    class BeamIntial
    {
        public static FilteredElementCollector BeamCollector(Document doc)
        {
            FilteredElementCollector beamCollector = new FilteredElementCollector(doc);
            beamCollector.OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilyInstance));
            TaskDialog.Show("REVIT", beamCollector.Count().ToString());
            return beamCollector;
        }
        public static void BeamElementExcelPara(Document doc, FilteredElementCollector beamCollector,
System.Data.DataTable dt)
        {
            foreach (Element ele in beamCollector)
            {
                FamilyInstance beam = ele as FamilyInstance;
                FamilySymbol familySymbol =
beam.Document.GetElement(beam.GetTypeId()) as FamilySymbol;
                if (beam != null)
                {
                    string Olev = null;
                    string material = null;
                    double length = 0.0;
                    double orOffset = 0.0;
                    double endOffset = 0.0;
                    double bottomHeight = 0.0;
                    double topHeight = 0.0;
                    string y_ = null;
                    double y_o = 0.0;
                    string yz_ = null;
                    string z_ = null;
                    double z_o = 0.0;
                    string use = null;
                    foreach (Parameter param in beam.Parameters)
                    {
                        InternalDefinition definition = null;
                        definition = param.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                        if (BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM == definition.BuiltInParameter)
                        {
                            Olev = param.AsValueString();
                        }
                        if (BuiltInParameter.STRUCTURAL_MATERIAL_PARAM == definition.BuiltInParameter)
                        {
                            material = param.AsValueString();
                        }
                        if (BuiltInParameter.INSTANCE_LENGTH_PARAM == definition.BuiltInParameter)
                        {
                            length = FeetTomm(param.AsDouble());
                        }
                        if (BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION == definition.BuiltInParameter)
                        {
                            orOffset = FeetTomm(param.AsDouble());
                        }
                        if (BuiltInParameter.STRUCTURAL_BEAM_END1_ELEVATION == definition.BuiltInParameter)
                        {
                            endOffset = FeetTomm(param.AsDouble());
                        }
                        if (BuiltInParameter.STRUCTURAL_ELEVATION_AT_BOTTOM == definition.BuiltInParameter)
                        {
                            bottomHeight = FeetTomm(param.AsDouble());
                        }
                        if (BuiltInParameter.STRUCTURAL_ELEVATION_AT_TOP == definition.BuiltInParameter)
                        {
                            topHeight = FeetTomm(param.AsDouble());
                        }
                        if (BuiltInParameter.Y_JUSTIFICATION == definition.BuiltInParameter)
                        {
                            y_ = param.AsValueString();
                        }
                        if (BuiltInParameter.Y_OFFSET_VALUE == definition.BuiltInParameter)
                        {
                            y_o = FeetTomm(param.AsDouble());
                        }
                        if (BuiltInParameter.YZ_JUSTIFICATION == definition.BuiltInParameter)
                        {
                            yz_ = param.AsValueString();
                        }
                        if (BuiltInParameter.Z_JUSTIFICATION == definition.BuiltInParameter)
                        {
                            z_ = param.AsValueString();
                        }
                        if (BuiltInParameter.Z_OFFSET_VALUE == definition.BuiltInParameter)
                        {
                            z_o = FeetTomm(param.AsDouble());
                        }
                        if (BuiltInParameter.INSTANCE_STRUCT_USAGE_PARAM == definition.BuiltInParameter)
                        {
                            use = param.AsValueString();
                        }
                    }
                    string area = null;
                    foreach (Parameter para in familySymbol.ParametersMap)
                    {
                        InternalDefinition definition = para.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                        if (BuiltInParameter.STRUCTURAL_SECTION_AREA == definition.BuiltInParameter)
                        {
                            area = para.AsValueString();
                        }
                    }
                    CreateBeamExcelRow(dt, beam.Symbol.FamilyName,
                        Olev, material, length, orOffset, endOffset, bottomHeight, topHeight,
                          y_, y_o, yz_, z_, z_o, use, area
                    );
                }
            }
        }
        private static void CreateBeamExcelRow(DataTable dt,
            string beamName, string Olev, string material,
     double length,
     double orOffset, double endOffset,
     double bottomHeight, double topHeight,
     string y_, double y_o,
     string yz_, string z_, double z_o, string use, string area
   )
        {
            DataRow dr = dt.NewRow();
            dr["梁名称"] = beamName;
            dr["参照标高"] = Olev;
            dr["结构材质"] = material;
            dr["长度"] = length;
            dr["起点标高偏移"] = orOffset;
            dr["终点标高偏移"] = endOffset;
            dr["底部高程"] = bottomHeight;
            dr["顶部高程"] = topHeight;
            dr["Y轴对正"] = y_;
            dr["Y轴偏移量"] = y_o;
            dr["YZ轴对正"] = yz_;
            dr["Z轴对正"] = z_;
            dr["Z轴偏移量"] = z_o;
            dr["结构用途"] = use;
            dr["截面面积"] = area;
            dt.Rows.Add(dr);
            dr = null;
        }
        public static System.Data.DataTable CreateBeamExcelTitle()
        {
            System.Data.DataTable dt = new System.Data.DataTable("梁信息表");
            DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true; dc.AutoIncrementSeed = 1; dc.AutoIncrementStep = 1; dc.AllowDBNull = false; dt.Columns.Add(new DataColumn("梁名称", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("参照标高", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("结构材质", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("长度", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("起点标高偏移", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("终点标高偏移", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("底部高程", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("顶部高程", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("Y轴对正", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Y轴偏移量", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("YZ轴对正", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Z轴对正", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Z轴偏移量", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("结构用途", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("截面面积", Type.GetType("System.String")));
            return dt;
        }
        public static void BeamElementMySQLPara(Document doc, FilteredElementCollector beamCollector,
System.Data.DataTable dt)
        {
            foreach (Element ele in beamCollector)
            {
                FamilyInstance beam = ele as FamilyInstance;
                FamilySymbol familySymbol =
beam.Document.GetElement(beam.GetTypeId()) as FamilySymbol;
                if (beam != null)
                {
                    string Olev = null;
                    string material = null;
                    double length = 0.0;
                    double orOffset = 0.0;
                    double endOffset = 0.0;
                    double bottomHeight = 0.0;
                    double topHeight = 0.0;
                    string y_duizheng = null;
                    double y_offset = 0.0;
                    string yz_duizheng = null;
                    string z_duizheng = null;
                    double z_offset = 0.0;
                    string use = null;
                    foreach (Parameter param in beam.Parameters)
                    {
                        InternalDefinition definition = null;
                        definition = param.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                        if (BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM == definition.BuiltInParameter)
                        {
                            Olev = param.AsValueString();
                        }
                        if (BuiltInParameter.STRUCTURAL_MATERIAL_PARAM == definition.BuiltInParameter)
                        {
                            material = param.AsValueString();
                        }
                        if (BuiltInParameter.INSTANCE_LENGTH_PARAM == definition.BuiltInParameter)
                        {
                            length = FeetTomm(param.AsDouble());
                        }
                        if (BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION == definition.BuiltInParameter)
                        {
                            orOffset = param.AsDouble();
                        }
                        if (BuiltInParameter.STRUCTURAL_BEAM_END1_ELEVATION == definition.BuiltInParameter)
                        {
                            endOffset = param.AsDouble();
                        }
                        if (BuiltInParameter.STRUCTURAL_ELEVATION_AT_BOTTOM == definition.BuiltInParameter)
                        {
                            bottomHeight = param.AsDouble();
                        }
                        if (BuiltInParameter.STRUCTURAL_ELEVATION_AT_TOP == definition.BuiltInParameter)
                        {
                            topHeight = param.AsDouble();
                        }
                        if (BuiltInParameter.Y_JUSTIFICATION == definition.BuiltInParameter)
                        {
                            y_duizheng = param.AsValueString();
                        }
                        if (BuiltInParameter.Y_OFFSET_VALUE == definition.BuiltInParameter)
                        {
                            y_offset = FeetTomm(param.AsDouble());
                        }
                        if (BuiltInParameter.YZ_JUSTIFICATION == definition.BuiltInParameter)
                        {
                            yz_duizheng = param.AsValueString();
                        }
                        if (BuiltInParameter.Z_JUSTIFICATION == definition.BuiltInParameter)
                        {
                            z_duizheng = param.AsValueString();
                        }
                        if (BuiltInParameter.Z_OFFSET_VALUE == definition.BuiltInParameter)
                        {
                            z_offset = FeetTomm(param.AsDouble());
                        }
                        if (BuiltInParameter.INSTANCE_STRUCT_USAGE_PARAM == definition.BuiltInParameter)
                        {
                            use = param.AsValueString();
                        }
                    }
                    CreateBeamMySQLRow(dt, beam.Symbol.FamilyName, Olev,
material, length,
y_duizheng, y_offset, yz_duizheng, z_duizheng, z_offset
);
                }
            }
        }
        private static void CreateBeamMySQLRow(DataTable dt, string beamName, string Olev,
string material, double length,
string y_duizheng, double y_offset,
string yz_duizheng, string z_duizheng, double z_offset
)
        {
            DataRow dr = dt.NewRow();
            dr["beamName"] = beamName;
            dr["Olev"] = Olev;
            dr["material"] = material;
            dr["length"] = length;
            dr["y_"] = y_duizheng;
            dr["y_o"] = y_offset;
            dr["yz_"] = yz_duizheng;
            dr["z_"] = z_duizheng;
            dr["z_o"] = z_offset;
            dt.Rows.Add(dr);
            dr = null;
        }
        public static System.Data.DataTable CreateBeamMySQLTitle()
        {
            System.Data.DataTable dt = new System.Data.DataTable("梁信息表");
            DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true; dc.AutoIncrementSeed = 1; dc.AutoIncrementStep = 1; dc.AllowDBNull = false; dt.Columns.Add(new DataColumn("beamName", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Olev", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("material", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("length", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("y_", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("y_o", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("yz_", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("z_", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("z_o", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("use", Type.GetType("System.String")));
            return dt;
        }
        public static double FeetTomm(double val)
        {
            return val * 304.8;
        }
    }
}

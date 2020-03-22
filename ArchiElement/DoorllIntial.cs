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
    class DoorIntial
    {
        public static FilteredElementCollector DoorCollector(Document doc)
        {
            FilteredElementCollector doorCollector = new FilteredElementCollector(doc);
            doorCollector.OfCategory(BuiltInCategory.OST_Doors).OfClass(typeof(FamilyInstance));
            TaskDialog.Show("REVIT", doorCollector.Count().ToString());
            return doorCollector;
        }
        public static void DoorElementExcelPara(Document doc, FilteredElementCollector doorCollector,
System.Data.DataTable dt)
        {
            foreach (Element ele in doorCollector)
            {

                FamilyInstance door = ele as FamilyInstance;

                Level level = level = doc.GetElement(ele.LevelId) as Level;

                FamilySymbol familySymbol =
    door.Document.GetElement(door.GetTypeId()) as FamilySymbol;

                if (door != null)
                {

                    string bottomheight = null;

                    string width = null;

                    string height = null;

                    string kjcha = null;

                    string drcha = null;

                    foreach (Parameter param in door.Parameters)
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
                        if (definition.Name == "框架材质")
                        {
                            kjcha = para.AsValueString();
                        }
                        if (definition.Name == "门材质")
                        {
                            drcha = para.AsValueString();
                        }
                    }

                    CreateDoorExcelRow(dt, door.Symbol.FamilyName,
                        level.Name, drcha, kjcha, width, bottomheight, height);
                }
            }
        }
        public static void DoorElementMySQLPara(Document doc, FilteredElementCollector doorCollector,
            System.Data.DataTable dt)
        {
            foreach (Element ele in doorCollector)
            {
                FamilyInstance door = ele as FamilyInstance;
                Level level = level = doc.GetElement(ele.LevelId) as Level;
                FamilySymbol familySymbol =
    door.Document.GetElement(door.GetTypeId()) as FamilySymbol;
                if (door != null)
                {
                    string bottomheight = null;
                    string width = null;
                    string height = null;
                    string kjcha = null;
                    string drcha = null;
                    foreach (Parameter param in door.Parameters)
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
                        if (definition.Name == "框架材质")
                        {
                            kjcha = para.AsValueString();
                        }
                        if (definition.Name == "门材质")
                        {
                            drcha = para.AsValueString();
                        }
                    }
                    CreateDoorMySQLRow(dt, door.Symbol.FamilyName,
                        level.Name, drcha, kjcha, width, bottomheight, height);
                }
            }
        }
        public static void CreateDoorExcelRow(System.Data.DataTable dt, string doortype, string levelName,
string drma, string kjma, string width, string bottomheight, string height)
        {
            DataRow dr = dt.NewRow();
            dr["门类型"] = doortype;
            dr["标高"] = levelName;
            dr["门材质"] = drma;
            dr["框架材质"] = kjma;
            dr["宽度"] = width;
            dr["底高度"] = bottomheight;
            dr["高度"] = height;
            dt.Rows.Add(dr);
            dr = null;
        }
        public static System.Data.DataTable CreateDoorExcelTitle()
        {
            System.Data.DataTable dt = new System.Data.DataTable("门信息表");
            DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true; dc.AutoIncrementSeed = 1; dc.AutoIncrementStep = 1; dc.AllowDBNull = false; dt.Columns.Add(new DataColumn("门类型", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("标高", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("门材质", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("框架材质", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("底高度", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("宽度", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("高度", Type.GetType("System.String")));
            return dt;
        }
        public static void CreateDoorMySQLRow(System.Data.DataTable dt, string doortype, string levelName,
string drma, string kjma, string width, string bottomheight, string height)
        {
            DataRow dr = dt.NewRow();
            dr["doortype"] = doortype;
            dr["levelName"] = levelName;
            dr["drma"] = drma;
            dr["kjma"] = kjma;
            dr["width"] = width;
            dr["bottomheight"] = bottomheight;
            dr["height"] = height;
            dt.Rows.Add(dr);
            dr = null;
        }
        public static System.Data.DataTable CreateDoorMySQLTitle()
        {
            System.Data.DataTable dt = new System.Data.DataTable("门信息表");
            DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true; dc.AutoIncrementSeed = 1; dc.AutoIncrementStep = 1; dc.AllowDBNull = false; dt.Columns.Add(new DataColumn("doortype", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("levelName", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("drma", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("kjma", Type.GetType("System.String")));
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

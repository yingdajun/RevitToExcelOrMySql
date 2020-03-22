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
using System.Data;
namespace ArchiElement
{
    class FloorIntial
    {
       public static FilteredElementCollector FloorCollector(Document doc) {
            FilteredElementCollector floorCollector = new FilteredElementCollector(doc);
            floorCollector.OfCategory(BuiltInCategory.OST_Floors).OfClass(typeof(Floor));
            return floorCollector;
        }
        public static void FloorElementExcelPara(Document doc, FilteredElementCollector floorCollector, System.Data.DataTable dt)
        {
            foreach (Element ele in floorCollector)
            {
                Floor floor = ele as Floor;
                Level level = doc.GetElement(floor.LevelId) as Level;
                if (floor != null)
                {
                    double thick = 0.0;
                    double area = 0.0;
                    double offset = 0.0;
                    foreach (Parameter param in floor.Parameters)
                    {
                      InternalDefinition definition = param.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                                                if (BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM == definition.BuiltInParameter)
                        {
                            thick = param.AsDouble();
                        }
                                                if (BuiltInParameter.HOST_AREA_COMPUTED == definition.BuiltInParameter)
                        {
                            area = param.AsDouble();
                        }
                                                if (BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM == definition.BuiltInParameter)
                        {
                            offset = param.AsDouble();
                        }
                    }
                    FloorIntial.
                    CreateFloorExcelRow(dt, floor.FloorType.Name, level.Name, FeetTomm(thick), FeetTomm(area)
                        , FeetTomm(offset));
                }
            }
        }
        public static void FloorElementMySQLPara(Document doc, FilteredElementCollector floorCollector, System.Data.DataTable dt)
        {
            foreach (Element ele in floorCollector)
            {
                Floor floor = ele as Floor;
                Level level = doc.GetElement(floor.LevelId) as Level;
                if (floor != null)
                {
                    double thick = 0.0;
                    double area = 0.0;
                    double offset = 0.0;
                    foreach (Parameter param in floor.Parameters)
                    {
                                                InternalDefinition definition = param.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                                                if (BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM == definition.BuiltInParameter)
                        {
                            thick = param.AsDouble();
                        }
                                                if (BuiltInParameter.HOST_AREA_COMPUTED == definition.BuiltInParameter)
                        {
                            area = param.AsDouble();
                        }
                                                if (BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM == definition.BuiltInParameter)
                        {
                            offset = param.AsDouble();
                        }
                    }
                    FloorIntial.
                    CreateFloorMySQLRow(dt, floor.FloorType.Name, level.Name, FeetTomm(thick), FeetTomm(area)
                        , FeetTomm(offset));
                }
            }
        }
                                                                                public static void CreateFloorExcelRow(System.Data.DataTable dt, string floorType, string levelName,
            double thick, double area, double offset)
        {
            DataRow dr = dt.NewRow();
            dr["楼板类型"] = floorType;
            dr["标高"] = levelName;
            dr["厚度"] = thick;
            dr["面积"] = area;
            dr["自标高的偏移"] = offset;
            dt.Rows.Add(dr);
            dr = null;
        }
                                        public static System.Data.DataTable CreateFloorExcelTitle()
        {
                        System.Data.DataTable dt =
                new System.Data.DataTable("楼板信息表");
                        DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true;               dc.AutoIncrementSeed = 1;              dc.AutoIncrementStep = 1;              dc.AllowDBNull = false;                            dt.Columns.Add(new DataColumn("楼板类型", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("标高", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("厚度", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("面积", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("自标高的偏移", Type.GetType("System.Decimal")));
            return dt;
        }
                                                                                public static void CreateFloorMySQLRow(System.Data.DataTable dt, string floorType, string levelName,
            double thick, double area, double offset)
        {
            DataRow dr = dt.NewRow();
            dr["floorType"] = floorType;
            dr["levelName"] = levelName;
            dr["thick"] = thick;
            dr["area"] = area;
            dr["offset"] = offset;
            dt.Rows.Add(dr);
            dr = null;
        }
                                        public static System.Data.DataTable CreateFloorMySQLTitle()
        {
                        System.Data.DataTable dt =
                new System.Data.DataTable("楼板信息表");
                        DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true;               dc.AutoIncrementSeed = 1;              dc.AutoIncrementStep = 1;              dc.AllowDBNull = false;                            dt.Columns.Add(new DataColumn("floorType", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("levelName", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("thick", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("area", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("offset", Type.GetType("System.Decimal")));
            return dt;
        }
                                                        public static double FeetTomm(double val)
        {
            return val * 304.8;
        }
                                                                                public static int Round(double number)
        {
            int count = 0;              string numstr = number.ToString();             string[] numsplit = numstr.Split('.');             int resultValue = Int32.Parse(numsplit[0]);
            int decimalValue = Int32.Parse(numsplit[1].Substring(0, 1));
            int len = numsplit[1].Length;                          if (len == 1)
            {
                if (decimalValue >= 5)
                {
                    resultValue += 1;                  }
            }
            else
            {
                                if (decimalValue <= 3)
                {
                    return resultValue;
                }
                else
                {
                    int[] numarray = new int[len];
                                        foreach (char chars in numsplit[1])
                    {
                        numarray[count] = Int32.Parse(chars.ToString());
                        count++;
                    }
                    for (int i = len - 1; i > 0; i--)
                    {
                                                if (numarray[i] >= 5)
                        {
                            numarray[i - 1] += 1;
                        }
                    }
                                        if (numarray[0] >= 5)
                    {
                        resultValue += 1;
                    }
                }
            }
            return resultValue;
        }
    }
}

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
    class WallIntial
    {
                                                public static FilteredElementCollector WallCollector(Document doc) {
            FilteredElementCollector wallCollector = new FilteredElementCollector(doc);
            wallCollector.OfCategory(BuiltInCategory.OST_Walls).OfClass(typeof(Wall));
            return wallCollector;
        }
        public static void WallElementExcelPara(Document doc, FilteredElementCollector wallCollector, System.Data.DataTable dt)
        {
            foreach (Element ele in wallCollector)
            {
                Wall wall = ele as Wall;
                Level level = doc.GetElement(wall.LevelId) as Level;
                if (wall != null)
                {
                    double height = 0.0;
                    foreach (Autodesk.Revit.DB.Parameter param in wall.Parameters)
                    {
                                                InternalDefinition definition = param.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                                                if (BuiltInParameter.WALL_USER_HEIGHT_PARAM == definition.BuiltInParameter)
                        {
                            height = param.AsDouble();
                        }
                    }
                                        CreateWallExcelRow(dt, wall.WallType.Name, level.Name, FeetTomm(wall.Width), FeetTomm(height));
                }
                            }
        }
        public static void WallElementMySQLPara(Document doc, FilteredElementCollector wallCollector, System.Data.DataTable dt)
        {
            foreach (Element ele in wallCollector)
            {
                Wall wall = ele as Wall;
                Level level = doc.GetElement(wall.LevelId) as Level;
                if (wall != null)
                {
                    double height = 0.0;
                    foreach (Autodesk.Revit.DB.Parameter param in wall.Parameters)
                    {
                                                InternalDefinition definition = param.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                                                if (BuiltInParameter.WALL_USER_HEIGHT_PARAM == definition.BuiltInParameter)
                        {
                            height = param.AsDouble();
                        }
                    }
                                        CreateWallMySQLRow(dt, wall.WallType.Name, level.Name, FeetTomm(wall.Width), FeetTomm(height));
                }
                            }
        }
                                                                        public static void CreateWallExcelRow(System.Data.DataTable dt, string walltype, string levelName,
            double width, double height)
        {
            DataRow dr = dt.NewRow();
            dr["墙体类型"] = walltype;
            dr["标高"] = levelName;
            dr["宽度"] = width;
            dr["高度"] = height;
            dt.Rows.Add(dr);
            dr = null;
        }
                                        public static System.Data.DataTable CreateWallExcelTitle()
        {
                        System.Data.DataTable dt = new System.Data.DataTable("墙的信息表");
                        DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true;               dc.AutoIncrementSeed = 1;              dc.AutoIncrementStep = 1;              dc.AllowDBNull = false;                            dt.Columns.Add(new DataColumn("墙体类型", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("标高", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("宽度", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("高度", Type.GetType("System.Decimal")));
            return dt;
        }
                                                                        public static void CreateWallMySQLRow(System.Data.DataTable dt, string walltype, string levelName,
            double width, double height)
        {
            DataRow dr = dt.NewRow();
            dr["walltype"] = walltype;
            dr["levelName"] = levelName;
            dr["width"] = width;
            dr["height"] = height;
            dt.Rows.Add(dr);
            dr = null;
        }
                                        public static System.Data.DataTable CreateWallMySQLTitle()
        {
                        System.Data.DataTable dt = new System.Data.DataTable("墙的信息表");
                        DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true;               dc.AutoIncrementSeed = 1;              dc.AutoIncrementStep = 1;              dc.AllowDBNull = false;                            dt.Columns.Add(new DataColumn("walltype", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("levelName", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("width", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("height", Type.GetType("System.Decimal")));
            return dt;
        }
                                                        public static double FeetTomm(double val)
        {
            return val * 304.8;
        }
    }
}

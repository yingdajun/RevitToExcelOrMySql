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
    class RoomIntial
    {
        public static FilteredElementCollector RoomCollector(Document doc)
        {
            FilteredElementCollector floorCollector = new FilteredElementCollector(doc);
            floorCollector.OfCategory(BuiltInCategory.OST_Rooms).OfClass(typeof(SpatialElement));
            return floorCollector;
        }
        public static void RoomElementExcelPara(Document doc, FilteredElementCollector roomCollector, DataTable dt)
        {
            foreach (Element ele in roomCollector)
            {
                Room room = ele as Room;
                Level level = doc.GetElement(room.LevelId) as Level;
                if (room != null)
                {
                    string number = null;
                    string name = null;
                    double area = 0.0;
                    double tpoffset = 0.0;
                    double bottomoffset = 0.0;
                    foreach (Parameter param in room.Parameters)
                    {
                        InternalDefinition definition = param.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                        if (BuiltInParameter.ROOM_NUMBER == definition.BuiltInParameter)
                        {
                            number = param.AsString();
                        }
                        if (BuiltInParameter.ROOM_NAME == definition.BuiltInParameter)
                        {
                            name = param.AsString();
                        }
                        if (BuiltInParameter.ROOM_AREA == definition.BuiltInParameter)
                        {
                            area = param.AsDouble();
                        }
                        if (BuiltInParameter.ROOM_LOWER_OFFSET == definition.BuiltInParameter)
                        {
                            bottomoffset = param.AsDouble();
                        }
                        if (BuiltInParameter.ROOM_UPPER_OFFSET == definition.BuiltInParameter)
                        {
                            tpoffset = param.AsDouble();
                        }
                    }
                    //怎么 又是这里有问题MMP
                    RoomIntial.
                    CreateRoomExcelRow(dt, number, name, level.Name, FeetTomm(bottomoffset), FeetTomm(area)
                        , FeetTomm(tpoffset));
                }
            }
        }
        public static void RoomElementMySQLPara(Document doc, FilteredElementCollector roomCollector, DataTable dt)
        {
            foreach (Element ele in roomCollector)
            {
                Room room = ele as Room;
                Level level = doc.GetElement(room.LevelId) as Level;
                if (room != null)
                {
                    string number = null;
                    string name = null;
                    double area = 0.0;
                    double tpoffset = 0.0;
                    double bottomoffset = 0.0;
                    foreach (Parameter param in room.Parameters)
                    {
                        InternalDefinition definition = param.Definition as InternalDefinition;
                        if (null == definition)
                            continue;
                        if (BuiltInParameter.ROOM_NUMBER == definition.BuiltInParameter)
                        {
                            number = param.AsString();
                        }
                        if (BuiltInParameter.ROOM_NAME == definition.BuiltInParameter)
                        {
                            name = param.AsString();
                        }
                        if (BuiltInParameter.ROOM_AREA == definition.BuiltInParameter)
                        {
                            area = param.AsDouble();
                        }
                        if (BuiltInParameter.ROOM_LOWER_OFFSET == definition.BuiltInParameter)
                        {
                            bottomoffset = param.AsDouble();
                        }
                        if (BuiltInParameter.ROOM_UPPER_OFFSET == definition.BuiltInParameter)
                        {
                            tpoffset = param.AsDouble();
                        }
                    }
                    RoomIntial.
                    CreateRoomMySQLRow(dt, number, name, level.Name, FeetTomm(bottomoffset), FeetTomm(area)
                        , FeetTomm(tpoffset));
                }
            }
        }
        public static void CreateRoomExcelRow(System.Data.DataTable dt, string number, string name,
string levelName, double bottomOffset, double area, double topoffset)
        {
            DataRow dr = dt.NewRow();
            dr["编号"] = number;
            dr["房间名"] = name;
            dr["标高"] = levelName;
            dr["底部偏移"] = bottomOffset;
            dr["面积"] = area;
            dr["高度偏移"] = topoffset;
            dt.Rows.Add(dr);
            dr = null;
        }
        public static System.Data.DataTable CreateRoomExcelTitle()
        {
            System.Data.DataTable dt =
    new System.Data.DataTable("房间信息表");
            DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true; dc.AutoIncrementSeed = 1; dc.AutoIncrementStep = 1; dc.AllowDBNull = false; dt.Columns.Add(new DataColumn("编号", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("房间名", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("标高", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("底部偏移", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("面积", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("高度偏移", Type.GetType("System.Decimal")));
            return dt;
        }




        public static void CreateRoomMySQLRow(System.Data.DataTable dt, string number, string name,
string levelName, double bottomOffset, double area, double topoffset)
        {
            DataRow dr = dt.NewRow();
            dr["number"] = number;
            dr["name"] = name;
            dr["levelName"] = levelName;
            dr["bottomOffset"] = bottomOffset;
            dr["area"] = area;
            dr["topoffset"] = topoffset;
            dt.Rows.Add(dr);
            dr = null;
        }
        public static DataTable CreateRoomMySQLTitle()
        {
            DataTable dt = new DataTable("房间信息表");
            DataColumn dc = dt.Columns.Add("id", Type.GetType("System.Int32"));
            dc.AutoIncrement = true; dc.AutoIncrementSeed = 1; dc.AutoIncrementStep = 1; dc.AllowDBNull = false; dt.Columns.Add(new DataColumn("number", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("name", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("levelName", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("bottomOffset", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("area", Type.GetType("System.Decimal")));
            dt.Columns.Add(new DataColumn("topoffset", Type.GetType("System.Decimal")));
            return dt;
        }
        public static double FeetTomm(double val)
        {
            return val * 304.8;
        }
        public static int Round(double number)
        {
            int count = 0; string numstr = number.ToString(); string[] numsplit = numstr.Split('.'); int resultValue = Int32.Parse(numsplit[0]);
            int decimalValue = Int32.Parse(numsplit[1].Substring(0, 1));
            int len = numsplit[1].Length; if (len == 1)
            {
                if (decimalValue >= 5)
                {
                    resultValue += 1;
                }
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

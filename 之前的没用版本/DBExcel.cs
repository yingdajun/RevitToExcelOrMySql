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
/// <summary>
/// 这个数据库是用来导入到EXCEL文件的
/// </summary>
namespace ArchiElement
{
    [Transaction(TransactionMode.Manual)]
    class DBExcel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //throw new NotImplementedException();
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            //获取所有墙体的墙体类型
            //FilteredElementCollector wallCollector = new FilteredElementCollector(doc);
            //wallCollector.OfCategory(BuiltInCategory.OST_Walls).OfClass(typeof(Wall));
            FilteredElementCollector roomCollector = new FilteredElementCollector(doc);

            roomCollector.OfCategory(BuiltInCategory.OST_Rooms).OfClass(typeof(SpatialElement));

            List<Element> roomList = roomCollector.ToList();

            TaskDialog.Show("room COUNT", roomList.Count().ToString());
            try {

                List<Element> elemList=ElementList(doc,"墙");
                List<Element> roomsList = WallList(doc);
                TaskDialog.Show("revit",roomsList.ToList().Count.ToString());
                //如果没有墙体直接退出
                if (!(roomsList.Count > 0))
                {
                    message = "模型没有房间";
                    return Result.Failed;
                }
                ////获取元素
                //List<string> createSetting=ShowDialog(roomsList);
                //if (!(createSetting.Count > 0))
                //{
                //    message = "没得参数值";
                //    return Result.Cancelled;
                //}

                //创建明细表---创建面积明细表测试
                Transaction t = new Transaction(doc,"创建明细表");
                FilteredElementCollector collector1 = new FilteredElementCollector(doc);
                collector1.OfCategory(BuiltInCategory.OST_AreaSchemes);
                //创建明细表
                //collector1.OfCategory(BuiltInCategory.OST_Schedules);
                ElementId areaSchemeId = collector1.FirstElementId();
                ViewSchedule Schedule = ViewSchedule.CreateSchedule(doc,new ElementId(BuiltInCategory.OST_Areas)
                    ,areaSchemeId);



                //添加字段到明细表
                //AddFieldToSchedule(Lis);
            } catch
            {

                return Result.Cancelled;
            }
            
            //RibbonWPF ribbonWPF = new RibbonWPF();
            //ribbonWPF.ShowDialog();

            //GetWallElement(doc);


            return Result.Succeeded;
        }


        /// <summary>
        /// 获取各类元素的List
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        private List<Element> ElementList(Document doc,string type)
        {
            List<Element> elemList = null;
            if (type=="墙体") {

            }

            FilteredElementCollector elemCollecotr = new FilteredElementCollector(doc);
            FilteredElementCollector wallCollector = new FilteredElementCollector(doc);
            wallCollector.OfCategory(BuiltInCategory.OST_Walls).OfClass(typeof(Wall));
            List<Element> wallList = wallCollector.ToList();
           



            //获取房间的元素

            FilteredElementCollector roomCollector = new FilteredElementCollector(doc);
            roomCollector.OfCategory(BuiltInCategory.OST_Rooms).OfClass(typeof(SpatialElement));
            List<Element> roomList = roomCollector.ToList();




            //获取地板元素

            FilteredElementCollector floorCollector = new FilteredElementCollector(doc);
            floorCollector.OfCategory(BuiltInCategory.OST_Floors).OfClass(typeof(Floor));
            List<Element> floorList = floorCollector.ToList();

            //获取视图的元素--三维的
            FilteredElementCollector viewCollector = new FilteredElementCollector(doc);
            viewCollector.OfCategory(BuiltInCategory.OST_Views).OfClass(typeof(View3D));
            List<Element> view3dList = viewCollector.ToList();

            FilteredElementCollector viewPlanCollector = new FilteredElementCollector(doc);
            viewPlanCollector.OfCategory(BuiltInCategory.OST_Views).OfClass(typeof(ViewPlan));
            List<Element> viewList = viewPlanCollector.ToList();

            //获取标高和轴网
            FilteredElementCollector levelCollector = new FilteredElementCollector(doc);
            levelCollector.OfCategory(BuiltInCategory.OST_Levels).OfClass(typeof(Level));
            List<Element> levelList = levelCollector.ToList();

            FilteredElementCollector gridCollector = new FilteredElementCollector(doc);
            gridCollector.OfCategory(BuiltInCategory.OST_Grids).OfClass(typeof(Grid));
            List<Element> gridList = gridCollector.ToList();

            //获取梁的元素
            //获取管道的元素
            //获取桥架的元素

         

            

           

            //FamilyInstance 元素
            //获取门的元素 
            FilteredElementCollector doorCollecotr = new FilteredElementCollector(doc);
            doorCollecotr.OfCategory(BuiltInCategory.OST_Doors).OfClass(typeof(FamilyInstance));
            //获取窗元素
            FilteredElementCollector windowsCollector = new FilteredElementCollector(doc);
            windowsCollector.OfCategory(BuiltInCategory.OST_Windows).OfClass(typeof(FamilyInstance));
            //获取柱子元素
            FilteredElementCollector columnCollector = new FilteredElementCollector(doc);
            columnCollector.OfCategory(BuiltInCategory.OST_Columns).OfClass(typeof(FamilyInstance));

            elemList = wallList;
            return elemList;


        }

        ///// <summary>
        ///// 弹出视图并且一共选择的
        ///// </summary>
        ///// <param name="roomList"></param>
        ///// <returns></returns>
        //private List<string> ShowDialog(List<Element> roomList)
        //{
        //    //初试化一个数组保存参数值
        //    List<string> value = new List<string>();

        //    //转化WallList近值内
        //    List<Room> roomsList = roomList.ConvertAll(x => x as Room);

        //    //初始化参数值
        //    List<string> parameterName = new List<string>();

        //    //获取参数对
        //    ParameterMap paraMap = roomList.First().ParametersMap;

        //    foreach (Parameter para in paraMap)
        //    {
        //        parameterName.Add(para.Definition.Name);
        //    }
        //    //传递到哪里去，传递到香港参数内
        //    //RibbonWPF ribbonWPF = new RibbonWPF(parameterName,roomsList);

        //    //if (ribbonWPF.ShowDialog()==true) {


        //    //    value.Add(ribbonWPF.wallParameterNameCombo.SelectedItem.ToString());

        //    //    value.Add(ribbonWPF.wallValueCombo.SelectedItem.ToString());

        //    //    return value;
        //    //}
        //    //return value;
        //}





        /// <summary>
        /// 获取Wall的元素
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        private List<Element> WallList(Document doc)
        {
            //throw new NotImplementedException();
            FilteredElementCollector wallCollector = new FilteredElementCollector(doc);
            wallCollector.OfCategory(BuiltInCategory.OST_Rooms).OfClass(typeof(SpatialElement));
            List<Element> wallList = wallCollector.ToList();
            return wallList;
        }



        private void GetWallElement(Document doc)
        {
            ElementClassFilter familyInstanceFilter = new ElementClassFilter(typeof(FamilyInstance));
            ElementCategoryFilter wallCategoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
            LogicalAndFilter wallFilter = new LogicalAndFilter(familyInstanceFilter,wallCategoryFilter);
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            IList<Element> walls = collector.WherePasses(wallFilter).ToElements();
            foreach (Element ele in walls) {
                Wall wall = ele as Wall;
                TaskDialog.Show("wall 特性",wall.WallType.Name.ToString());
            }
        }

        /// <summary>
        /// 获取墙体某些元素
        /// </summary>
        /// <param name="uidoc"></param>
        /// <param name="doc"></param>
        public void GetWallTaskDialog(UIDocument uidoc,Document doc) {

            Transaction t = new Transaction(doc, "1");
            t.Start();
            Reference refer = uidoc.Selection.PickObject(ObjectType.Element, "选择墙体");
            Wall wall = doc.GetElement(refer) as Wall;
            //可以能做的出来
            string[] s = {wall.Id.ToString(), wall.Width.ToString() , wall.WallType.ToString() , wall.Name.ToString()
            ,wall.LevelId.ToString() };
            for (int i = 0; i < s.Length; i++)
            {
                TaskDialog.Show("Revit", s[i]);
            }
            t.Commit();
        }
    }
}

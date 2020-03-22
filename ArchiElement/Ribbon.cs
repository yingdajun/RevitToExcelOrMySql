using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.IO;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
namespace ArchiElement
{
    class Ribbon : IExternalApplication
    {
        string AddinPath = typeof(Ribbon).Assembly.Location;
        public Result OnShutdown(UIControlledApplication application)
        {
            TaskDialog.Show("小树签结束", "小树签");
            return Result.Succeeded;
        }
        public Result OnStartup(UIControlledApplication application)
        {
            TaskDialog.Show("小树签成功使用", "小树签的功能是提取BIM模型信息");
            string tab = "小树签工具";
            application.CreateRibbonTab(tab);
            RibbonPanel panel = application.CreateRibbonPanel(tab, "导出BIM信息到EXCEL表格和MYSQL数据库中");
            AddSpliteButton(panel);
            return Result.Succeeded;
        }
        private void AddSpliteButton(RibbonPanel panel)
        {
            PushButtonData wallbtn = new PushButtonData("wallName", "墙信息导出", AddinPath, "ArchiElement.WallTest");
            wallbtn.LargeImage = new BitmapImage(new Uri(AddinPath.Replace("ArchiElement.dll", "wall.ico")));
            wallbtn.ToolTip = "将项目中墙参数信息导出";
            PushButtonData floorbtn = new PushButtonData("floorName", "楼板信息导出", AddinPath, "ArchiElement.FloorTest");
            floorbtn.LargeImage = new BitmapImage(new Uri(AddinPath.Replace("ArchiElement.dll", "floor.ico")));
            floorbtn.ToolTip = "将项目中楼板参数信息导出";
            PushButtonData doorbtn = new PushButtonData("doorName", "门信息导出", AddinPath, "ArchiElement.DoorTest");
            doorbtn.LargeImage = new BitmapImage(new Uri(AddinPath.Replace("ArchiElement.dll", "door.ico")));
            doorbtn.ToolTip = "将项目中门参数信息导出";
            PushButtonData windowbtn = new PushButtonData("windowName", "窗信息导出", AddinPath, "ArchiElement.WindowTest");
            windowbtn.LargeImage = new BitmapImage(new Uri(AddinPath.Replace("ArchiElement.dll", "window.ico")));
            windowbtn.ToolTip = "将项目中窗参数信息导出";
            PushButtonData roombtn = new PushButtonData("roomName", "房间信息导出", AddinPath, "ArchiElement.RoomTest");
            roombtn.LargeImage = new BitmapImage(new Uri(AddinPath.Replace("ArchiElement.dll", "room.ico")));
            roombtn.ToolTip = "将项目中房间参数信息导出";
            SplitButtonData sp1 = new SplitButtonData("建筑", "建筑");
            SplitButton sb1 = panel.AddItem(sp1) as SplitButton;
            sb1.AddPushButton(wallbtn);
            sb1.AddPushButton(floorbtn);
            sb1.AddPushButton(doorbtn);
            sb1.AddPushButton(windowbtn);
            sb1.AddPushButton(roombtn);
            panel.AddSeparator();
            PushButtonData columnbtn = new PushButtonData("columnName", "结构柱信息导出", AddinPath, "ArchiElement.ColumnTest");

            columnbtn.LargeImage = new BitmapImage(new Uri(AddinPath.Replace("ArchiElement.dll", "column.ico")));
            columnbtn.ToolTip = "将项目中结构柱参数信息导出";
            PushButtonData beambtn = new PushButtonData("beamName", "梁信息导出", AddinPath, "ArchiElement.BeamTest");
            beambtn.LargeImage = new BitmapImage(new Uri(AddinPath.Replace("ArchiElement.dll", "beam.ico")));
            beambtn.ToolTip = "将项目中梁参数信息导出";
            SplitButtonData sp2 = new SplitButtonData("结构", "结构");
            SplitButton sb2 = panel.AddItem(sp2) as SplitButton;
            sb2.AddPushButton(columnbtn);
            sb2.AddPushButton(beambtn);
            panel.AddSeparator();   
            PushButtonData ductbtn = new PushButtonData("ductName", "风管信息导出", AddinPath, "ArchiElement.DuctTest");
            ductbtn.LargeImage = new BitmapImage(new Uri(AddinPath.Replace("ArchiElement.dll", "duct.ico")));
            ductbtn.ToolTip = "将项目中风管参数信息导出";
            PushButtonData pipebtn = new PushButtonData("pipeguanName", "水管信息导出", AddinPath, "ArchiElement.PipeTest");
            pipebtn.LargeImage = new BitmapImage(new Uri(AddinPath.Replace("ArchiElement.dll", "pipe.ico")));
            pipebtn.ToolTip = "将项目中水管参数信息导出";
            PushButtonData qiaobtn = new PushButtonData("elqName", "桥架信息导出", AddinPath, "ArchiElement.EletricalQTest");
            qiaobtn.LargeImage = new BitmapImage(new Uri(AddinPath.Replace("ArchiElement.dll", "qiaojia.ico"))); 
            qiaobtn.ToolTip = "将项目中电缆桥架参数信息导出";
            SplitButtonData sp3 = new SplitButtonData("机电", "机电");
            SplitButton sb3 = panel.AddItem(sp3) as SplitButton;
            sb3.AddPushButton(ductbtn);
            sb3.AddPushButton(pipebtn);
            sb3.AddPushButton(qiaobtn);
        }
        public FileSystemInfo PickFolderInfo()
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择EXCEL文件所放置的位置";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                return new DirectoryInfo(dialog.SelectedPath);
            }
            return null;
        }
    }
}

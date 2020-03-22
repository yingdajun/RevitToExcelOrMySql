using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
namespace ArchiElement
{
    class Ribbon2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //throw new NotImplementedException();
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            //ElementExtract_WPF wpf = new ElementExtract_WPF();
            //RibbonWPF wpf = new RibbonWPF();


            ////判断是否选择了
            //if (batchWpf.FamilyPath == null)
            //{
            //    return Result.Cancelled;
            //}
            ////判断选择
            ////batchWpf.ShowDialog();
            //if (batchWpf.ShowDialog() == false)
            //{
            //    return Result.Cancelled;
            //}

            ////判断是否选择了
            //if (wpf.ExcelFilePath == null)
            //{
            //    return Result.Cancelled;
            //}
            ////判断选择
            ////batchWpf.ShowDialog();
            //if (wpf.ShowDialog() == false)
            //{
            //    return Result.Cancelled;
            //}

            return Result.Succeeded;

        }
    }
}

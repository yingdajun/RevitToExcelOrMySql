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
namespace ArchiElement
{
    [Transaction(TransactionMode.Manual)]
    class ParameterDemoTest : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //throw new NotImplementedException();
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            FilteredElementCollector doorCollector = new FilteredElementCollector(doc);

            doorCollector.OfCategory(BuiltInCategory.OST_Doors).OfClass(typeof(FamilyInstance));

            TaskDialog.Show("REVIT",doorCollector.Count().ToString());


            //获取元素
            Reference refer = uidoc.Selection.PickObject(ObjectType.Element);
            //Element ele = doc.GetElement(refer);

            FamilyInstance fsi = doc.GetElement(refer) as FamilyInstance;
            
            //门类型
            TaskDialog.Show("revit", fsi.Symbol.FamilyName);

 

            string bottomheight = null;
            foreach (Parameter param in fsi.Parameters) {
              

                //获取参数
                InternalDefinition definition = param.Definition as InternalDefinition;
                if (null == definition)
                    continue;
                //获取底高度
    
                if (BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM == definition.BuiltInParameter)
                {
                    //thickness = param.AsDouble();
                    bottomheight = param.AsValueString();

                }

               

            }



            FamilySymbol familySymbol =  

            fsi.Document.GetElement(fsi.GetTypeId()) as FamilySymbol;


   
           

            //门宽度
            string width = null;

            foreach (Parameter para in familySymbol.ParametersMap) {

                InternalDefinition definition = para.Definition as InternalDefinition;
                if (null == definition)
                    continue;
                if (BuiltInParameter.DOOR_WIDTH == definition.BuiltInParameter)
                {
 
                    width = para.AsValueString();

                }
            }
            TaskDialog.Show("宽度",width);



            //门高度
            string height = null;


            foreach (Parameter para in familySymbol.ParametersMap)
            {

                InternalDefinition definition = para.Definition as InternalDefinition;
                if (null == definition)
                    continue;
                if (BuiltInParameter.GENERIC_HEIGHT == definition.BuiltInParameter)
                {
                    //height = param.AsDouble();
                    //= para.AsValueString();
                    height = para.AsValueString();

                }
            }
            TaskDialog.Show("高度", height);

            ////框架材质
 
            string  kuangjia= null;


            foreach (Parameter para in familySymbol.ParametersMap)
            {

                InternalDefinition definition = para.Definition as InternalDefinition;
                if (null == definition)
                    continue;

                if (definition.Name== "框架材质")
                {
     
                    kuangjia = para.AsValueString();

                }
            }
            TaskDialog.Show("框架材质", kuangjia);


            ///门材质

            string doorchaizhi = null;


            foreach (Parameter para in familySymbol.ParametersMap)
            {

                InternalDefinition definition = para.Definition as InternalDefinition;
                if (null == definition)
                    continue;
  
                if (definition.Name == "门材质")
                {
      
                    doorchaizhi = para.AsValueString();

                }
            }
            TaskDialog.Show("门材质", doorchaizhi);

            //string kuangjiaChai = "";
            //foreach (Parameter para in familySymbol.Parameters)
            //{
            //    //DOOR_FRAME_MATERIAL
            //    //获取参数
            //    InternalDefinition definition = para.Definition as InternalDefinition;
            //    if (null == definition)
            //        continue;
            //    //获取框架材质
            //    if (BuiltInParameter.INVALID == definition.BuiltInParameter)
            //    {
            //        //height = param.AsDouble();
            //        kuangjiaChai = para.AsValueString();

            //    }
            //}

            //TaskDialog.Show("2", kuangjiaChai);



            TaskDialog.Show("1",bottomheight);

            ////楼板类型
            //TaskDialog.Show("revit", floor.FloorType.Name);
            ////标高
            //TaskDialog.Show("revit",level.Name);
            ////厚度
            //TaskDialog.Show("revit", FeetTomm(thickness).ToString());
            ////面积
            //TaskDialog.Show("revit",FeetTomm(area).ToString());
            ////偏移
            //TaskDialog.Show("revit",FeetTomm(offset).ToString());
            //获取楼板名称
            //

            //MessageBox.Show(strParamInfo);

            //TaskDialog.Show("revit",);


            //FindParameter(ele);

            return Result.Succeeded;
        }



        public void Get_FamilySymbol(Family family)//获取族类型
        {
            //TaskDialog.Show("1", "1");
            StringBuilder message = new StringBuilder("选择的族文件名称: " + family.Name + "\n " + "\n ");
            ISet<ElementId> familySymbolIds = family.GetFamilySymbolIds();
            if (familySymbolIds.Count == 0) { message.AppendLine("Contains no family symbols."); }
            else
            {
                message.AppendLine("文件中有以下类型 : ");
                foreach (ElementId id in familySymbolIds)
                {
                    FamilySymbol familySymbol = family.Document.GetElement(id) as FamilySymbol;
                    message.AppendLine("\nName: " + familySymbol.Name);
                    foreach (ElementId materialId in familySymbol.GetMaterialIds(false))
                    {
                        Material material = familySymbol.Document.GetElement(materialId) as Material;
                        message.AppendLine("\nMaterial : " + material.Name);
                    }
                }
            }
        }

        public static double FeetTomm(double val)
        {
            return val * 304.8;
        }

        private Parameter FindParameter(Element ele)
        {
            //throw new NotImplementedException();
            Floor floor = ele as Floor;
            Parameter foundParameter = null;
            foreach (Parameter parameter in ele.Parameters) {
                if (parameter.Definition.ParameterType==ParameterType.Length) {
                    foundParameter = parameter;
                    break;
                   
                }
            }
            return foundParameter;
        }
    }
}

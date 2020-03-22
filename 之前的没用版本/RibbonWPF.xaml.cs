using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
using System.IO;
using System.Windows.Forms;
namespace ArchiElement
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
   

    public partial class RibbonWPF : Window
    {
        /// <summary>
        /// 这里是界面
        /// </summary>
        //List<Room> roomsList;

        public FileSystemInfo ExcelFilePath { get; set; }
        public RibbonWPF()
        {
            InitializeComponent();
            //选择EXCEL文件
            ExcelFilePath = PickFolderInfo();

        }


        ///// <summary>
        ///// 获取参数的元素和对象
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void WallParameterNameCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{
        //    ////获取选择的参数名称
        //    //string paraName = wallParameterNameCombo.SelectedIndex.ToString();
        //    ////参数值列表
        //    //List<string> parameterValueList = new List<string>();
        //    ////加入全部生成按钮
        //    //parameterValueList.Add("无");
        //    ////遍历所有房间元素
        //    //foreach (Room room in roomsList)
        //    //{

        //    //    //获取元素的参数和对象进行遍历
        //    //    //元素对组列表
        //    //    ParameterMap paraMap = room.ParametersMap;
        //    //    //遍历元素对组
        //    //    foreach (Parameter para in paraMap)
        //    //    {
        //    //        //如果元素的定义
        //    //        if (para.Definition.Name == paraName)
        //    //        {
        //    //            //参数的HASVALUE值
        //    //            if (para.HasValue)
        //    //            {
        //    //                //初始化参数值
        //    //                string value;
        //    //                //如果参数化存储的值==存储的
        //    //                if (para.StorageType == StorageType.String)
        //    //                {
        //    //                    //元素值等于
        //    //                    value = para.AsString();
        //    //                }
        //    //                else
        //    //                {
        //    //                    //元素值等于
        //    //                    value = para.AsValueString();
        //    //                }
        //    //                //如果元素不包含元素，加入元素
        //    //                if (!parameterValueList.Contains(value))
        //    //                {
        //    //                    parameterValueList.Add(value);
        //    //                }
        //    //            }
        //    //        }
        //    //    }
        //    //}

        //    //wallValueCombo.ItemsSource = parameterValueList;
        //}

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //wallParameterNameCombo.SelectedIndex = 0;
            //wallValueCombo.SelectedIndex = 0;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ////判断命令
            //if (wallParameterNameCombo.SelectedItem != null
            //    && wallValueCombo.SelectedItem != null
            //  )
            //{
            //    //如果对话框有错误，关闭掉
            //    this.DialogResult = true;
            //    this.Close();
            //}
        }

        private void Path_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Outdata_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Outhtml_Click(object sender, RoutedEventArgs e)
        {

        }

        private FileSystemInfo PickFolderInfo()
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择EXCEL文件所在地址";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                return new DirectoryInfo(dialog.SelectedPath);
            }
            return null;
        }
    }
}

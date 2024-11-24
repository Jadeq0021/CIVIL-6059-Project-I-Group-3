using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;//类Application使用

namespace Task
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]
    public class Class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Application app = commandData.Application.Application;//引用:using Autodesk.Revit.ApplicationServices;
            //创建对话框
            TaskDialog mainDialog = new TaskDialog("isBIM模术师");//对话框的名称
            mainDialog.MainInstruction = "产品使用说明";//对话框标题
            mainDialog.MainContent = "isBIM模术师是基于Autodesk Revit软件的本地化功能插件集";//对话框的主要内容
            mainDialog.ExpandedContent = "可用于建筑、结构、水电以及暖通等专业中";//对话框的扩展内容
            //添加命令链接
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "查看当前Revit版本信息");
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "查看模术师产品信息");
            mainDialog.CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel;
            mainDialog.VerificationText = "不再显示该信息";
            //添加文字超链接
            mainDialog.FooterText = "<a href=\"http://www.bimcheng.com\">" + "点此处了解更多信息</a>";

            //显示对话框并取得返回值
            TaskDialogResult tResult = mainDialog.Show();

            //使用对话框返回结果
            if (TaskDialogResult.CommandLink1 == tResult)
            {
                //链接一的扩展的对话框内容
                TaskDialog dialog_CommandLink1 = new TaskDialog("版本信息");
                dialog_CommandLink1.MainInstruction = "版本名:" + app.VersionNumber + "\n" + "版本号:"
                    + app.VersionNumber;
                dialog_CommandLink1.Show();
            }
            else if (TaskDialogResult.CommandLink2 == tResult)
            {
                //链接二的扩展的对话框内容
                TaskDialog.Show("模术师产品介绍", "isBIM模术师是一个全过程、全专业的高效解决方案");
            }

            return Result.Succeeded;
        }
    }
}
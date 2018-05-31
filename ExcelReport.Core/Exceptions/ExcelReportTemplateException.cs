/*
 类：ExcelReportTemplateException
 描述：ExcelReport模板异常
 编 码 人：韩兆新 日期：2018年5月30日17:19:01
 修改记录：

*/
namespace NPOI.ExcelReport.Core.Exceptions
{
    internal class ExcelReportTemplateException : ExcelReportException
    {
        public ExcelReportTemplateException(string message)
            : base(message)
        {
        }
    }
}
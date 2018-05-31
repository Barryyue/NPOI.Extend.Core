/*
 类：ExcelReportException
 描述：ExcelReport异常
 编 码 人：韩兆新 日期：2018年5月30日17:19:01
 修改记录：

*/

using System;

namespace NPOI.ExcelReport.Core.Exceptions
{
    public class ExcelReportException : ApplicationException
    {
        public ExcelReportException(string message) : base(message)
        {
        }
    }
}
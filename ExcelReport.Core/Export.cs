/*
 类：Export
 描述：导出
 编 码 人：韩兆新 日期：2018年5月30日17:19:01
 修改记录：

*/

using NPOI.Extend.Core;
using NPOI.Extend.Core.Extend;
using NPOI.SS.UserModel;

namespace NPOI.ExcelReport.Core
{
    public static class Export
    {
        /// <summary>
        ///     导出格式化处理后的文件到二进制文件流
        /// </summary>
        /// <param name="templateFile"></param>
        /// <param name="sheetFormatters"></param>
        /// <returns></returns>
        public static byte[] ExportToBuffer(string templateFile, params SheetFormatter[] sheetFormatters)
        {
            IWorkbook workbook = NPOIHelper.LoadWorkbook(templateFile);
           
            foreach (SheetFormatter sheetFormatter in sheetFormatters)
            {
                sheetFormatter.Format(workbook);
            }
            return workbook.SaveToBuffer();
        }
    }
}
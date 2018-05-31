/*
 类：ElementFormatter
 描述：元素格式化器接口
 编 码 人：韩兆新  日期：2018年5月30日17:19:01
 修改记录：
 *2015年06月17日(幸运草)  修改方法SetCellValue(),补充对TypeCode.Decimal的处理;
 *2018年5月30日17:19:01（韩兆新） 修改自ExcelReport_v1.xx中的ElementFormatter。
*/
namespace NPOI.ExcelReport.Core.Formatters
{
    public abstract class ElementFormatter
    {
        /// <summary>
        ///     格式化
        /// </summary>
        /// <param name="sheetAdapter">Sheet适配器</param>
        public abstract void Format(SheetAdapter sheetAdapter);
    }
}
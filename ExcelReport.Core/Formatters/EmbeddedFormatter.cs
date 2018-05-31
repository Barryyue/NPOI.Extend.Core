/*
 类：EmbeddedFormatter
 描述：内嵌格式化器接口
 编 码 人：韩兆新  日期：2018年5月30日17:19:01
 修改记录：
 *
 
*/
namespace NPOI.ExcelReport.Core.Formatters
{
    public abstract class EmbeddedFormatter<TSource>
    {
        /// <summary>
        ///     格式化
        /// </summary>
        /// <param name="sheetAdapter">Sheet适配器</param>
        /// <param name="dataSource">数据源</param>
        public abstract void Format(SheetAdapter sheetAdapter, TSource dataSource);
    }
}
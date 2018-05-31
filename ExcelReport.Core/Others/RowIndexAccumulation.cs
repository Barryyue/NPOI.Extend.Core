/*
 类：RowIndexAccumulation
 描述：行标累加器
 编 码 人：韩兆新 日期：2018年5月30日17:19:01
 修改记录：
 * 

*/
namespace NPOI.ExcelReport.Core.Others
{
    public class RowIndexAccumulation : Accumulation
    {
        #region 1.0 获取当前行标

        /// <summary>
        ///     获取当前行标
        /// </summary>
        /// <param name="sourceRowIndex">源行标</param>
        /// <returns></returns>
        public int GetCurrentRowIndex(int sourceRowIndex)
        {
            return Value + sourceRowIndex;
        }

        #endregion
    }
}
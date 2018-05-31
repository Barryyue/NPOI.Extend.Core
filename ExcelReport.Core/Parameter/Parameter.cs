/*
 类：Parameter
 描述：参数信息
 编 码 人：韩兆新 日期：2018年5月30日17:19:01
 修改记录：
 * 2018年5月30日17:19:01  删除CellPoint，添加RowIndex、ColumnIndex。

*/

namespace NPOI.ExcelReport.Core.Parameter
{
    public class Parameter
    {
        public string Name { set; get; }
        public int RowIndex { set; get; }
        public int ColumnIndex { set; get; }
    }
}
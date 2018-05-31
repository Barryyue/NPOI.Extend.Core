/*
 结构：CellSpan
 描述：单元格跨度信息
 编 码 人：韩兆新 日期：2018年5月30日17:19:01
 修改记录：

*/

namespace NPOI.Extend.Core.Info
{
    public struct CellSpan
    {
        private int _colSpan;
        private int _rowSpan;

        public CellSpan(int rowSpan, int colSpan)
        {
            _rowSpan = rowSpan;
            _colSpan = colSpan;
        }

        public int RowSpan
        {
            get { return _rowSpan; }
            set { _rowSpan = value; }
        }

        public int ColSpan
        {
            get { return _colSpan; }
            set { _colSpan = value; }
        }
    }
}
﻿/*
 类：SheetFormatter
 描述：Sheet格式化器
 编 码 人：韩兆新 日期：2018年5月30日17:19:01
 修改记录：
 * 2018年5月30日17:19:01（韩兆新） 

*/

using System.Collections.Generic;
using NPOI.ExcelReport.Core.Extend;
using NPOI.ExcelReport.Core.Formatters;
using NPOI.SS.UserModel;

namespace NPOI.ExcelReport.Core
{
    public class SheetFormatter
    {
        #region 成员字段及属性
        private readonly IList<ElementFormatter> _formatterList;
        private readonly string _sheetName;

        public string SheetName
        {
            get { return _sheetName; }
        }

        public IList<ElementFormatter> FormatterList
        {
            get { return _formatterList; }
        }
        #endregion

        #region 0 构成函数
        /// <summary>
        /// 实例化Sheet格式化器
        /// </summary>
        /// <param name="sheetName">Sheet名字</param>
        /// <param name="formatters">格式化器集合</param>
        public SheetFormatter(string sheetName, params ElementFormatter[] formatters)
        {
            this._sheetName = sheetName;
            this._formatterList = new List<ElementFormatter>(formatters);
        }
        #endregion

        #region 1.0 格式化
        /// <summary>
        /// 格式化
        /// </summary>
        /// <param name="workbook">工作薄</param>
        public void Format(IWorkbook workbook)
        {
            ISheet sheet = workbook.GetSheet(SheetName);
            if (!sheet.IsNull() && !FormatterList.IsNullOrEmpty())
            {
                var sheetAdapter = new SheetAdapter(sheet);
                foreach (ElementFormatter formatter in FormatterList)
                {
                    formatter.Format(sheetAdapter);
                }
            }
        }
        #endregion
    }
}
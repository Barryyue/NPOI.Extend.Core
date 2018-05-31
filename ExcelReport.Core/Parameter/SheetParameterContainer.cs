/*
 类：SheetParameterContainer
 描述：工作表参数容器
 编 码 人：韩兆新 日期：2018年5月30日17:19:01
 修改记录：
 * 2018年5月30日17:19:01（韩兆新）  。

*/

using System.Collections.Generic;
using NPOI.ExcelReport.Core.Exceptions;

namespace NPOI.ExcelReport.Core.Parameter
{
    public class SheetParameterContainer
    {
        private List<Parameter> _parameterList = new List<Parameter>();
        public string SheetName { set; get; }

        public List<Parameter> ParameterList
        {
            set { _parameterList = value; }
            get { return _parameterList; }
        }

        /// <summary>
        ///     参数
        /// </summary>
        /// <param name="name">参数名</param>
        /// <returns>参数</returns>
        public Parameter this[string name]
        {
            get
            {
                foreach (Parameter parameter in _parameterList)
                {
                    if (parameter.Name.Equals(name))
                    {
                        return parameter;
                    }
                }
                throw new ExcelReportTemplateException("parameter is not exists");
            }
            set
            {
                bool isExist = false;
                foreach (Parameter parameter in _parameterList)
                {
                    if (parameter.Name.Equals(name))
                    {
                        isExist = true;
                        parameter.RowIndex = value.RowIndex;
                        parameter.ColumnIndex = value.ColumnIndex;
                    }
                }
                if (!isExist)
                {
                    _parameterList.Add(value);
                }
            }
        }
    }
}
using System;
using System.DrawingCore;
using System.Text.RegularExpressions;
using CoreExamples.Model;
using NPOI.ExcelReport.Core;
using NPOI.ExcelReport.Core.Extend;
using NPOI.ExcelReport.Core.Formatters.Complex;
using NPOI.ExcelReport.Core.Formatters.Simple;
using NPOI.ExcelReport.Core.Parameter;
using NPOI.Extend.Core;
using NPOI.SS.UserModel;

namespace CoreExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.DarkGreen;
            CreateExcel01();
            Console.WriteLine("1.0 success \r\n");

            CreateExcel02();
            Console.WriteLine("2.0 success \r\n");

            CreateExcel03();
            Console.WriteLine("3.0 success \r\n");

            CreateExcel04();
            
            Console.WriteLine("4.0 Success \r\n");

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("★ all success ★");

            Console.ReadKey();
        }

        /// <summary>
        /// 局部格式化器
        /// </summary>
        private static void CreateExcel01()
        {
            var temp = ($@"{AppDomain.CurrentDomain.BaseDirectory}1.0\");
            var excel = temp + "Template.xls";
            var xml = temp + "Template.xml";

            Parse(excel).Save(xml);

            var workbookParameterContainer = new WorkbookParameterContainer();
            workbookParameterContainer.Load(xml);
            var sheetParameterContainer = workbookParameterContainer["局部格式化器"];

            ExportHelper.ExportToLocal(excel, temp + "1.0demo.xls",
                new SheetFormatter("局部格式化器",
                    new PartFormatter(sheetParameterContainer["UserName"], "Jensen"),
                    new PartFormatter(sheetParameterContainer["GroupNo"], "116476496")
                )
            );
        }

        /// <summary>
        /// 单元格格式化器
        /// </summary>
        private static void CreateExcel02()
        {
            var temp = ($@"{AppDomain.CurrentDomain.BaseDirectory}2.0\");
            var excel = temp + "Template.xls";
            var xml = temp + "Template.xml";

            Parse(excel).Save(xml);

            var workbookParameterContainer = new WorkbookParameterContainer();
            workbookParameterContainer.Load(xml);
            var sheetParameterContainer = workbookParameterContainer["单元格格式化器"];

            ExportHelper.ExportToLocal(excel, temp + "2.0demo.xls",
                new SheetFormatter("单元格格式化器",
                    new CellFormatter(sheetParameterContainer["String"], "Hello World!"),
                    new CellFormatter(sheetParameterContainer["Boolean"], true),
                    new CellFormatter(sheetParameterContainer["DateTime"], DateTime.Now),
                    new CellFormatter(sheetParameterContainer["Double"], 3.14),
                    new CellFormatter(sheetParameterContainer["Image"], Image.FromFile(temp + "C#高级编程.jpg").ToBuffer())
                )
            );
        }

        /// <summary>
        /// 表格格式化器
        /// </summary>
        private static void CreateExcel03()
        {
            var temp = ($@"{AppDomain.CurrentDomain.BaseDirectory}3.0\");
            var excel = temp + "Template.xls";
            var xml = temp + "Template.xml";

            Parse(excel).Save(xml);

            var workbookParameterContainer = new WorkbookParameterContainer();
            workbookParameterContainer.Load(xml);
            SheetParameterContainer sheetParameterContainer = workbookParameterContainer["表格格式化器"];

            int num = 0;
            ExportHelper.ExportToLocal(excel, temp + "3.0demo.xls",
                new SheetFormatter("表格格式化器",
                    new TableFormatter<StudentInfo>(sheetParameterContainer["No"], StudentLogic.GetList(),
                        new CellFormatter<StudentInfo>(sheetParameterContainer["No"], t => num++),
                        new CellFormatter<StudentInfo>(sheetParameterContainer["Name"], t => t.Name),
                        new CellFormatter<StudentInfo>(sheetParameterContainer["Gender"], t => t.Gender ? "男" : "女"),
                        new CellFormatter<StudentInfo>(sheetParameterContainer["Class"], t => t.Class),
                        new CellFormatter<StudentInfo>(sheetParameterContainer["RecordNo"], t => t.RecordNo),
                        new CellFormatter<StudentInfo>(sheetParameterContainer["Phone"], t => t.Phone),
                        new CellFormatter<StudentInfo>(sheetParameterContainer["Email"], t => t.Email)
                    )
                )
            );
        }

        /// <summary>
        /// 重复单元格格式化器
        /// </summary>
        private static void CreateExcel04()
        {
            var temp = ($@"{AppDomain.CurrentDomain.BaseDirectory}4.0\");
            var excel = temp + "Template.xls";
            var xml = temp + "Template.xml";

            Parse(excel).Save(xml);

            var workbookParameterContainer = new WorkbookParameterContainer();
            workbookParameterContainer.Load(xml);
            SheetParameterContainer sheetParameterContainer = workbookParameterContainer["重复单元格式化器"];

            ExportHelper.ExportToLocal(excel, temp + "4.0demo.xls",
                new SheetFormatter("重复单元格式化器",
                    new RepeaterFormatter<StudentInfo>(sheetParameterContainer["rptStudentInfo_Start"],
                        sheetParameterContainer["rptStudentInfo_End"], StudentLogic.GetList(),
                        new CellFormatter<StudentInfo>(sheetParameterContainer["Name"], t => t.Name),
                        new CellFormatter<StudentInfo>(sheetParameterContainer["Gender"], t => t.Gender ? "男" : "女"),
                        new CellFormatter<StudentInfo>(sheetParameterContainer["Class"], t => t.Class),
                        new CellFormatter<StudentInfo>(sheetParameterContainer["RecordNo"], t => t.RecordNo),
                        new CellFormatter<StudentInfo>(sheetParameterContainer["Phone"], t => t.Phone),
                        new CellFormatter<StudentInfo>(sheetParameterContainer["Email"], t => t.Email)
                    )
                )
            );
        }


        public static WorkbookParameterContainer Parse(string templatePath)
        {
            var workbookParameterContainer = new WorkbookParameterContainer();
            IWorkbook workbook = NPOIHelper.LoadWorkbook(templatePath);
            foreach (ISheet sheet in workbook)
            {
                workbookParameterContainer[sheet.SheetName] = new SheetParameterContainer
                {
                    SheetName = sheet.SheetName
                };
                foreach (IRow row in sheet)
                {
                    foreach (ICell cell in row.Cells)
                    {
                        if (cell.CellType.Equals(CellType.String))
                        {
                            MatchCollection matches = new Regex(@"(?<=\$\[)([\w]*)(?=\])").Matches(cell.StringCellValue);
                            foreach (Match match in matches)
                            {
                                workbookParameterContainer[sheet.SheetName][match.Value] = new Parameter
                                {
                                    Name = match.Value,
                                    RowIndex = cell.RowIndex,
                                    ColumnIndex = cell.ColumnIndex
                                };
                            }
                        }
                    }
                }
            }
            return workbookParameterContainer;
        }
    }
}


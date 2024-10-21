using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace ExcelMerge
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // 直接双击运行
            if (args.Length == 0)
            {
                Console.WriteLine("合并Excel文件，直接将文件拖拽到此程序即可！\r\n按ESC键关闭此窗口，按Enter键合并本目录下所有 xls xlsx 文件，若当前目录下没有Excel文件，则按config.yaml中的配置处理!");
                ConsoleKeyInfo key = Console.ReadKey();
                if (key.Key == ConsoleKey.Escape)
                {
                    return;
                }
                else
                {
                    if (key.Key == ConsoleKey.Enter)
                    {
                        string CurrentDirectory = Environment.CurrentDirectory;
                        DirectoryInfo di = new DirectoryInfo(CurrentDirectory);

                        List<SrcExcel> lstSrcExcel = new List<SrcExcel>();
                        //列程序目录下所有的Excel文件 xls xlsx
                        foreach (FileInfo fi in di.GetFiles("*.*"))
                        {
                            bool ExcelTag = false;

                            string ExtName = fi.Extension.ToLower();
                            switch (ExtName)
                            {
                                case ".xls":
                                    ExcelTag = true; break;
                                case ".xlsx":
                                    ExcelTag = true; break;
                                default:
                                    ExcelTag = false;
                                    break;
                            }

                            if (ExcelTag)
                            {
                                SrcExcel excel = new SrcExcel();
                                excel.FileFullName = fi.FullName;
                                lstSrcExcel.Add(excel);
                            }
                        }

                        if (lstSrcExcel.Count > 1)
                        {
                            MergeExcel(lstSrcExcel);
                            Console.WriteLine("按Enter键退出...");
                            Console.ReadLine();
                        }
                        else // 如果目录下没有Excel文件，则读取yaml文件作为配置参数
                        {
                            string configYaml = Path.Combine(CurrentDirectory, "config.yaml");
                            if (File.Exists(configYaml))
                            {
                                var yaml = File.ReadAllText(configYaml);

                                var deserializer = new DeserializerBuilder().Build();
                                try
                                {
                                    var config = deserializer.Deserialize<Config>(yaml);
                                    MergeExcel(config);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex.Message);
                                    Console.ReadLine();
                                }
                            }
                        }
                    }

                }
            }
            else
            {
                // 一个文件
                if (args.Length == 1)
                {
                    // 不处理
                }
                else  // 多个文件
                {
                    List<SrcExcel> lstSrcExcel = new List<SrcExcel>();
                    foreach (string src in args)
                    {
                        if (File.Exists(src))
                        {
                            string ExtName = Path.GetExtension(src).ToLower();
                            if (ExtName.Equals(".xlsx") || ExtName.Equals(".xls"))
                            {
                                SrcExcel excel = new SrcExcel();
                                excel.FileFullName = src;
                                lstSrcExcel.Add(excel);
                            }
                        }
                    }

                    if (lstSrcExcel.Count > 1)
                    {
                        MergeExcel(lstSrcExcel);
                        Console.WriteLine("按Enter键退出...");
                        Console.ReadLine();
                    }
                }
            }
        }


        // 合并Excel文件，默认不进行自动列宽操作
        private static void MergeExcel(List<SrcExcel> lstSrcExcel, bool AutoSizeColumn = false)
        {
            IWorkbook outputWorkbook = new XSSFWorkbook();

            foreach (SrcExcel excel in lstSrcExcel)
            {
                Console.WriteLine($"正在合并 {excel.FileFullName} ...");

                string FileFullName = excel.FileFullName;
                if (File.Exists(FileFullName))
                {
                    using (var stream = new FileStream(FileFullName, FileMode.Open, FileAccess.Read))
                    {
                        IWorkbook workbook;

                        string ExtName = Path.GetExtension(FileFullName).ToLower();
                        //XSSFWorkbook 用于处理 .xlsx 文件，HSSFWorkbook 用于处理 .xls 文件
                        if (ExtName.Equals(".xlsx"))
                        {
                            workbook = new XSSFWorkbook(stream);
                        }
                        else
                        {
                            workbook = new HSSFWorkbook(stream);
                        }

                        // 同步输出表和输入表的Sheet 以及每个Sheet的首行 作为标题列
                        SyncSheet(outputWorkbook, workbook);


                        int SheetNumber = workbook.NumberOfSheets;
                        for (int i = 0; i < SheetNumber; i++)
                        {
                            string SheetName = workbook.GetSheetName(i);
                            ISheet outputSheet = outputWorkbook.GetSheet(SheetName);

                            ISheet inputSheet = workbook.GetSheet(SheetName);

                            // 跳过首行标题行 同步其他所有行
                            for (int row = 1; row <= inputSheet.LastRowNum; row++)
                            {
                                IRow inputRow = inputSheet.GetRow(row);
                                IRow outputRow = outputSheet.CreateRow(outputSheet.LastRowNum + 1);
                                if (inputRow != null)
                                {
                                    for (int col = 0; col < inputRow.LastCellNum; col++)
                                    {
                                        ICell cell = inputRow.GetCell(col);
                                        if (cell != null)
                                        {
                                            ICell newCell = outputRow.CreateCell(col);

                                            // 复制单元格类型
                                            newCell.SetCellType(cell.CellType);

                                            // 复制单元格值
                                            newCell.SetCellValue(cell.ToString());

                                            // 复制单元格样式
                                            ICellStyle newStyle = outputWorkbook.CreateCellStyle();
                                            newStyle.CloneStyleFrom(cell.CellStyle); // 克隆样式
                                            newCell.CellStyle = newStyle; // 应用样式

                                            // 复制单元格备注
                                            newCell.CellComment = cell.CellComment;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }


            // 自动列宽
            if (AutoSizeColumn)
            {
                Console.WriteLine($"正在对合并后的Excel文件进行自动列宽操作...");

                int OutputSheetNumber = outputWorkbook.NumberOfSheets;
                for (int i = 0; i < OutputSheetNumber; i++)
                {
                    ISheet OutputSheet = outputWorkbook.GetSheetAt(i);
                    if (OutputSheet != null)
                    {
                        IRow OutputRow = OutputSheet.GetRow(0);
                        if (OutputRow != null)
                        {
                            for (int j = 0; j < OutputRow.LastCellNum; j++)
                            {
                                OutputSheet.AutoSizeColumn(j);
                            }
                        }
                    }
                }
            }

            // 生成合并的Excel文件
            var outputFile = Path.Combine(Environment.CurrentDirectory, $"合并结果({DateTime.Now.Hour}-{DateTime.Now.Minute}-{DateTime.Now.Second}).xlsx");
            using (var fileStream = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
            {
                outputWorkbook.Write(fileStream);
            }
            Console.WriteLine($"合并完成 {outputFile}!");
        }

        // 按yaml配置文件处理合并Excel
        private static void MergeExcel(Config config)
        {
            List<SrcExcel> lstSrcExcel = new List<SrcExcel>();
            foreach (string src in config.Excels)
            {
                if (File.Exists(src))
                {
                    string ExtName = Path.GetExtension(src).ToLower();
                    if (ExtName.Equals(".xlsx") || ExtName.Equals(".xls"))
                    {
                        SrcExcel excel = new SrcExcel();
                        excel.FileFullName = src;
                        lstSrcExcel.Add(excel);
                    }
                }
            }

            if (lstSrcExcel.Count > 1)
            {
                MergeExcel(lstSrcExcel, config.AutoSizeColumn);
                Console.WriteLine("按换行键退出...");
                Console.ReadLine();
            }
        }


        // 同步输出表和输入表的Sheet 以及每个Sheet的首行 作为标题列
        private static void SyncSheet(IWorkbook outputWorkbook, IWorkbook workbook)
        {
            int SheetNumber = workbook.NumberOfSheets;
            for (int i = 0; i < SheetNumber; i++)
            {
                string SheetName = workbook.GetSheetName(i);
                if (outputWorkbook.GetSheet(SheetName) == null)
                {
                    ISheet outputSheet = outputWorkbook.CreateSheet(SheetName);

                    ISheet inputSheet = workbook.GetSheet(SheetName);

                    IRow inputRow = inputSheet.GetRow(0);
                    IRow outputRow = outputSheet.CreateRow(0);

                    for (int col = 0; col < inputRow.LastCellNum; col++)
                    {
                        ICell cell = inputRow.GetCell(col);
                        if (cell != null)
                        {
                            ICell newCell = outputRow.CreateCell(col);
                            newCell.SetCellValue(cell.ToString());
                        }
                    }
                }
            }
        }

    }


    public class SrcExcel
    {
        public string FileFullName;
    }

    public class Config
    {
        public bool AutoSizeColumn { get; set; }
        public List<string> Excels { get; set; }
    }


}

using CommandLine;
using CommandLine.Text;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeExcel
{
    class Options
    {

        [Option('F', "folder", Required = true, HelpText = @"需要合并文件所在的文件夹路径")]
        public string Folder { get; set; }

        [Option('o', "output", Required = true, HelpText = @"合并后的文件路径")]
        public string OutputPath { get; set; }

        [Option('h', "headers", Required = true, HelpText = @"需要合并的字段")]
        public IEnumerable<string> Headers { get; set; }

        [Option('f', "files", Required = false, HelpText = @"需要合并的文件")]
        public IEnumerable<string> InputFiles { get; set; }
        
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                CommandLine.Parser.Default.ParseArguments<Options>(args).WithParsed<Options>(opts => main(opts));
                Console.ReadLine();
            }catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            /*
            DataTable dt = new DataTable();
            string[] files = new string[] { @"..\..\merge1.xlsx", @"..\..\merge2.xlsx" };
            for (int i = 0; i < files.Length; i++)
            {
                MergeData(files[i], dt);
            }
            ExportDataTableToExcel(dt, @"..\..\result.xlsx");
            */
        }

        private static void main(Options options)
        {
            string[] files;
            DataTable dt = new DataTable();

            if (options.InputFiles.Count() > 0)
            {
                files = options.InputFiles.ToArray();
            }else
            {
                files = Directory.GetFiles(options.Folder, "*.xlsx", SearchOption.AllDirectories);
            }

            for (int i = 0; i < files.Length; i++)
            {
                Console.WriteLine("{0}. 合并：{1}",(i + 1), files[i]);
                MergeData(files[i], dt, options.Headers.ToArray());
            }

            Console.WriteLine("生成：{0}", options.OutputPath);
            ExportDataTableToExcel(dt, options.OutputPath);
        }

        private static void MergeData(string path, DataTable dt, string[] headers)
        {
            // write data in workbook from xls document.
            XSSFWorkbook workbook = new XSSFWorkbook(path);
            // read the current table data
            XSSFSheet sheet = (XSSFSheet)workbook.GetSheetAt(0);
            // read the current row data
            XSSFRow headerRow = (XSSFRow)sheet.GetRow(0);
            // LastCellNum is the number of cells of current rows
            int cellCount = headerRow.LastCellNum;

            if (dt.Rows.Count == 0)
            {

                // build header for there is no data after the first implementation
                for (int i = 0; i < headers.Length; i++)
                {
                    
                    // get data as the column header of DataTable
                    DataColumn column = new DataColumn(headers[i]);

                    dt.Columns.Add(column);
                }
            }
            else
            {

                // TODO: check if the subsequent sheet corresponds
            }
            // LastRowNum is the number of rows of current table
            int rowCount = sheet.LastRowNum + 1;
            var colsIndexMapper = GetColsIndexMapper(headers, sheet);
            var colsIndexArray = colsIndexMapper.Values.ToArray();
            for (int i = (sheet.FirstRowNum + 1); i < rowCount; i++)
            {
                XSSFRow row = (XSSFRow)sheet.GetRow(i);

                bool isEmptyRow = !row.Any(e => colsIndexArray.Contains(e.ColumnIndex) && !string.IsNullOrEmpty(e.StringCellValue));

                if (isEmptyRow) continue;

                DataRow dataRow = dt.NewRow();

                for (var j = 0; j < headers.Length; ++j)
                {
                    var titleIndex = colsIndexMapper[headers[j]];

                    if(row.GetCell(titleIndex) != null && !string.IsNullOrEmpty(row.GetCell(titleIndex).ToString()))
                    {
                        dataRow[titleIndex] = row.GetCell(titleIndex).ToString();

                    }
                }
                /*
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                            // get data and convert them into character string type, then save them into the rows of datatable
                            dataRow[j] = row.GetCell(j).ToString();


                    }
                 * */
                dt.Rows.Add(dataRow);
            }
            workbook = null;
            sheet = null;
        }

        public static Dictionary<string,int> GetColsIndexMapper(string[] titles, XSSFSheet xSSFSheet)
        {
            var dict = new Dictionary<string, int>();
            XSSFRow headerRow = (XSSFRow)xSSFSheet.GetRow(xSSFSheet.FirstRowNum);

            foreach(var title in titles)
            {
                for(var i = headerRow.FirstCellNum; i < headerRow.LastCellNum;++i)
                {
                    var headerName = headerRow.GetCell(i).ToString().Trim();

                    if (title == headerName)
                        dict[title] = i;
                    
                }

            }

            return dict;
        }

        public static void ExportDataTableToExcel(DataTable dtSource, string strFileName)
        {

            XSSFWorkbook workbook = new XSSFWorkbook();
            
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("mySheet");

            XSSFRow dataRow = (XSSFRow)sheet.CreateRow(0);
            foreach (DataColumn column in dtSource.Columns)
            {
                // create the cells in the first row, and add data into these cells circularly
                dataRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);

            }
            //create rows on the basis of data from datatable(not including table header), and add data into cells in every row
            for (int i = 0; i < dtSource.Rows.Count; i++)
            {
                dataRow = (XSSFRow)sheet.CreateRow(i + 1);
                for (int j = 0; j < dtSource.Columns.Count; j++)
                {
                    dataRow.CreateCell(j).SetCellValue(dtSource.Rows[i][j].ToString());
                }
            }
            using (MemoryStream ms = new MemoryStream())
            {
                using (FileStream fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                {

                    workbook.Write(fs);
                }
            }
        }
    }
}

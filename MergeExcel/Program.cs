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
        [Option('f', "files", Required = false, HelpText = "需要合并的文件")]
        public IEnumerable<string> InputFiles { get; set; }

        [Option('F', "folder", Required = true, HelpText = "需要合并的文件夹路径")]
        public string Folder { get; set; }

        [Option('t', "titles", Required = true, HelpText = "需要合并的字段")]
        public IEnumerable<string> Titles { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            CommandLine.Parser.Default.ParseArguments<Options>(args).WithParsed<Options>(opts => main(opts));

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
            if (options.InputFiles.Count() > 0)
            {
                files = options.InputFiles.ToArray();
            }else if (!string.IsNullOrEmpty(options.Folder))
            {
                files = Directory.GetFiles(options.Folder, "*.xlsx");
            }


        }

        private static void MergeData(string path, DataTable dt, string[] titles)
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
                for (int i = 0; i < titles.Length; i++)
                {
                    
                    // get data as the column header of DataTable
                    DataColumn column = new DataColumn(titles[i]);

                    dt.Columns.Add(column);
                }
            }
            else
            {

                // TODO: check if the subsequent sheet corresponds
            }
            // LastRowNum is the number of rows of current table
            int rowCount = sheet.LastRowNum + 1;
            for (int i = (sheet.FirstRowNum + 1); i < rowCount; i++)
            {
                XSSFRow row = (XSSFRow)sheet.GetRow(i);
                DataRow dataRow = dt.NewRow();
                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (row.GetCell(j) != null)
                        // get data and convert them into character string type, then save them into the rows of datatable
                        dataRow[j] = row.GetCell(j).ToString();


                }
                dt.Rows.Add(dataRow);
            }
            workbook = null;
            sheet = null;
        }

        public Dictionary<string,int> GetColsIndexMapper(string[] titles, XSSFSheet xSSFSheet)
        {
            return null;
        }

        public static void ExportDataTableToExcel(DataTable dtSource, string strFileName)
        {
            // create workbook
            XSSFWorkbook workbook = new XSSFWorkbook();
            // the table named mySheet
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("mySheet");
            // create the first row
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

                    workbook.Write(fs);// write mySheet table in xls document and save it
                }
            }
        }
    }
}

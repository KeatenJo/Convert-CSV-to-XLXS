using System;
using System.IO;
using OfficeOpenXml;


namespace csvConvertxls
{
    class Program
    {
        static void Main(string[] args)
        {
            if(args.Length > 0)
            {
                foreach (string arg in args)
                {
                    Convert(arg);
                }
            } else
            {
                Console.WriteLine("Please Supply the name of your csv file.");
                Console.WriteLine("Ex.  ./csvConvertxls foo.txt bar.txt");
            }
            
        }


        public static void Convert(string name)
        {
            string csvFileName = name;
            string[] temp_name = name.Split(".");
            string excelFileName = temp_name[0] + ".xlsx";

            string worksheetsName = temp_name[0];

            bool firstRowIsHeader = false;

            var format = new ExcelTextFormat();
            format.Delimiter = ',';
            format.EOL = "\r";              // DEFAULT IS "\r\n";

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFileName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetsName);
                worksheet.Cells["A1"].LoadFromText(new FileInfo(csvFileName), format, OfficeOpenXml.Table.TableStyles.Medium27, firstRowIsHeader);
                package.Save();
            }

            Console.WriteLine("Finished converting " + csvFileName + " to " + excelFileName + ".");
        }
    }
}

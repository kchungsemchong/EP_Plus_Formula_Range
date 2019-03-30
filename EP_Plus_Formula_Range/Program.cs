using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EP_Plus_Formula_Range
{
	class Program
	{
		static void Main(string[] args)
		{
			string formula = File.ReadLines(@"C:\Testing\ExcelFormula\formula.txt").First();
			string csv_file_path = @"C:\Testing\ExcelFormula\test_formula_1.csv";

			// Creating the excel package
			using(ExcelPackage package = new ExcelPackage())
			{
				//set the formatting options
				ExcelTextFormat format = new ExcelTextFormat();
				format.Delimiter = ';';
				format.Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString());
				format.Culture.DateTimeFormat.ShortDatePattern = "dd-mm-yyyy";
				format.Encoding = new UTF8Encoding();

				//read the CSV file from disk
				FileInfo csv_file = new FileInfo(csv_file_path);


				// Creating the worksheet
				ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet 1");
				worksheet.Cells["A1"].LoadFromText(csv_file, format);

				// Using range in epplus
				// Worksheet.Cells[From row, from column, to row, to column]
				// Adding the range

				//Defining the tables parameters
				int firstRow = 1;
				int lastRow = worksheet.Dimension.End.Row;
				int firstColumn = 1;
				int lastColumn = worksheet.Dimension.End.Column;
				ExcelRange range_for_table = worksheet.Cells[firstRow, firstColumn, lastRow, lastColumn];
				string tableName = "Table1";

				//Ading a table to a Range
				ExcelTable tab = worksheet.Tables.Add(range_for_table, tableName);

				////Adding another range
				int first_row = 2;
				int last_row = 5;
				int first_column = 4;
				int last_column = 4;


				//var range_for_formula = worksheet.Cells[first_row, first_column, last_row, last_column];
				var range_for_formula = worksheet.Cells["D2:D6"];
				range_for_formula.Formula = "=SUM([Col1],[Col2])";
				range_for_formula.Calculate();


				//Formating the table style
				tab.TableStyle = TableStyles.Light8;

				if (File.Exists(@"C:\Testing\ExcelFormula\result.xlsx"))
				{
					File.Delete(@"C:\Testing\ExcelFormula\result.xlsx");
				}

				FileInfo result = new FileInfo(@"C:\Testing\ExcelFormula\result.xlsx");
				package.SaveAs(result);


				Console.WriteLine("Result saved!. Press any key to continue");
				Console.ReadKey();
			}



		}
	}
}

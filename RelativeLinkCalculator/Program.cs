using System;
using System.Text;
using NanoXLSX;
using RelativeLinkCalculator.Extensions;

namespace RelativeLinkCalculator // Note: actual namespace depends on the project name.
{
	internal class Program
	{
		const int kCompanyCodeColumn = 0;			// 公司代碼
		const int kDateColumn = 1;					// 年月日
		const int kStockHolderNameColumn = 2;		// 持股人姓名
		const int kHolderNameColumn = 3;			// 持有人姓名
		const int kPositionColumn = 4;				// 身分別
		const int kRelativeOrSummaryColumn = 5;		// 親屬/總說明

		static void Main(string[] args)
		{
			System.Console.OutputEncoding = Encoding.Unicode;
			string fileName = "Input//2022橡膠工業.xlsx";
			Workbook workbook = Workbook.Load(fileName);
			parseWorkbook(workbook);
		}

		private static void parseWorkbook(Workbook workbook)
		{
			Worksheet sheet = workbook.Worksheets[0];

			IReadOnlyList<Cell> columnA = sheet.GetColumn("A");
			int numberOfEntries = columnA.Count;

			System.Console.Write($"Number Of Entries: {numberOfEntries}\n\n");

			string currentCompanyName = "";
			int currentCompanyTotalPosition = 0;

			// We start with index 1, because the first row is always the title of each column.
			for (int rowIndex = 1; rowIndex < numberOfEntries; rowIndex++)
			{
				Address companyCodeAddress = new Address(column: kCompanyCodeColumn, row: rowIndex);
				Address stockHolderNameAddress = new Address(column: kStockHolderNameColumn, row: rowIndex);
				Address holderNameAddress = new Address(column: kHolderNameColumn, row: rowIndex);

				bool isSumupRow = !sheet.HasCell(holderNameAddress);
				if (isSumupRow)
				{
					// If holder name cell is empty (or cannot be found), it means the row is a sum-up row (i.e., not an legal individual);
					// therefore we want to skip these rows.
					continue;
				}

				if (!sheet.TryGetCell(new Address(column: kCompanyCodeColumn, row: rowIndex), out Cell companyCodeCell))
				{
					System.Console.WriteLine($"Row {rowIndex} has no company name (公司代碼), skip this row.");
					continue;
				}

				string companyName = companyCodeCell.Value.ToString();
				if (companyName.IsNullOrEmpty())
				{
					System.Console.WriteLine($"Row {rowIndex} has a null/empty company name (公司代碼), skip this row.");
					continue;
				}

				bool isNewCompany = companyName != currentCompanyName;
				if (isNewCompany) 
				{
					if (!currentCompanyName.IsNullOrEmpty()) 
					{
						System.Console.WriteLine($"{currentCompanyName}: {currentCompanyTotalPosition} 個職位");
					}

					currentCompanyName = companyName;
					currentCompanyTotalPosition = 0;
				}

				currentCompanyTotalPosition++;
			}

			Console.Read();
		}
	}
}
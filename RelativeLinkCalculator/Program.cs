using System;
using System.Text;
using System.Text.RegularExpressions;
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

		static Regex kIndependentDirectorPositionPattern = new Regex(@"獨立董事");
		static Regex kBoardPositionPattern = new Regex(@"董事");

		struct CompanyData
		{
			public string Name;

			public int RelativeLinkCount;
			public int TotalPositionCount;
			public int IndependentDirectorCount;

			public int BoardRelativeLinkCount;
			public int BoardPositionCount;

			/// <summary>
			/// Holder (member) name -> A list of the names of the member relatives
			/// </summary>
			public Dictionary<string, HashSet<string>> RelativeMap;
			public HashSet<string> BoardAndIndependentDirectors;
		}

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

			List<CompanyData> compantDataCollection = new List<CompanyData>();
			CompanyData currentCompanyData = new CompanyData();

			// We start with index 1, because the first row is always the title of each column.
			for (int rowIndex = 1; rowIndex < numberOfEntries; rowIndex++)
			{
				Address companyCodeAddress = new Address(column: kCompanyCodeColumn, row: rowIndex);
				Address holderNameAddress = new Address(column: kHolderNameColumn, row: rowIndex);
				Address positionAddress = new Address(column: kPositionColumn, row: rowIndex);

				bool isSumupRow = !sheet.HasCell(holderNameAddress);
				if (isSumupRow)
				{
					// If holder name cell cannot be found), it means the row is a sum-up row (i.e., not an legal individual);
					// therefore we want to skip these rows.
					continue;
				}

				Cell holderNameCell = sheet.GetCell(holderNameAddress);
				bool hasHolderName = !holderNameCell.Value.ToString().IsNullOrEmpty();
				if (!hasHolderName)
				{
					// If holder name cell cannot be found), it means the row is a sum-up row (i.e., not an legal individual);
					// therefore we want to skip these rows.
					continue;
				}

				if (!sheet.TryGetCell(companyCodeAddress, out Cell companyCodeCell))
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

				bool isNewCompany = companyName != currentCompanyData.Name;
				if (isNewCompany) 
				{
					if (!currentCompanyData.Name.IsNullOrEmpty()) 
					{
						System.Console.WriteLine($"{currentCompanyData.Name}: " +
							$"\t{currentCompanyData.TotalPositionCount} 個職位, " +
							$"\t{currentCompanyData.IndependentDirectorCount} 個獨立董事, " +
							$"\t{currentCompanyData.BoardPositionCount} 個董事");
					}
					compantDataCollection.Add(currentCompanyData);

					// Reset data counters
					currentCompanyData = new CompanyData();
					currentCompanyData.Name = companyName;
					currentCompanyData.RelativeMap = new Dictionary<string, HashSet<string>>();
					currentCompanyData.BoardAndIndependentDirectors = new HashSet<string>();
				}

				// Do various counting.

				currentCompanyData.TotalPositionCount++;

				bool hasPositionCell = sheet.TryGetCell(positionAddress, out Cell positionCell) && !positionCell.Value.ToString().IsNullOrEmpty();
				if (hasPositionCell) 
				{
					string positionName = positionCell.Value.ToString();

					bool isIndependentDirector = kIndependentDirectorPositionPattern.IsMatch(positionName);
					if (isIndependentDirector) 
					{
						currentCompanyData.IndependentDirectorCount++;
					}

					bool isBoardMember = kBoardPositionPattern.IsMatch(positionName);
					if (isBoardMember)
					{
						currentCompanyData.BoardPositionCount++;
					}
				}
			}

			Console.Read();
		}
	}
}
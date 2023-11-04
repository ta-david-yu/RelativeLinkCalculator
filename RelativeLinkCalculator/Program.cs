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

		static readonly Regex kIndependentDirectorPositionPattern = new Regex(@"獨立董事");
		static readonly Regex kBoardPositionPattern = new Regex(@"董事");
		static readonly Regex kNotRelativeColumnPattern = new Regex(@"集團");

		class CompanyData
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
			public Dictionary<string, HashSet<string>> RelativeMap = new Dictionary<string, HashSet<string>>();
			public HashSet<string> BoardAndIndependentDirectors = new HashSet<string>();
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
				Address relativeAddress = new Address(column: kRelativeOrSummaryColumn, row: rowIndex);

				bool isSumUpRow = !sheet.TryGetCell(holderNameAddress, out Cell holderNameCell) || holderNameCell.DataType != Cell.CellType.STRING;
				if (isSumUpRow)
				{
					// If holder name cell cannot be found), it means the row is a sum-up row (i.e., not an legal individual);
					// therefore we want to skip these rows.
					continue;
				}

				string holderName = holderNameCell.Value.ToString();
				if (holderName.IsNullOrEmpty())
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
						compantDataCollection.Add(currentCompanyData);
					}

					// Reset data counters
					currentCompanyData = new CompanyData();
					currentCompanyData.Name = companyName;

					System.Console.WriteLine($"-- {companyName} --");
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
						currentCompanyData.BoardAndIndependentDirectors.Add(holderName.Trim());
					}
				}

				bool hasRelativeOrSummaryCell = sheet.TryGetCell(relativeAddress, out Cell relativeCell);
				if (!hasRelativeOrSummaryCell) 
				{
					continue;
				}

				bool isRelativeCell = relativeCell.DataType == Cell.CellType.STRING && !kNotRelativeColumnPattern.IsMatch(relativeCell.Value.ToString());
				if (isRelativeCell) 
				{
					string allRelativeNameInput = relativeCell.Value.ToString();

					if (!currentCompanyData.RelativeMap.ContainsKey(holderName)) 
					{
						currentCompanyData.RelativeMap.Add(holderName, new HashSet<string>());
					}

					string[] relativeNamesWithSuffix = allRelativeNameInput.Split(',');
					foreach (string relativeNameWithSuffix in relativeNamesWithSuffix) 
					{
						bool hasValidName = relativeNameWithSuffix.TryExtractRelativeName(out string relativeName);
						if (!hasValidName) 
						{
							continue;
						}

						currentCompanyData.RelativeMap[holderName].Add(relativeName);

						bool hasLinkBeenRecorded = currentCompanyData.RelativeMap.ContainsKey(relativeName) && 
												   currentCompanyData.RelativeMap[relativeName].Contains(holderName);

						if (hasLinkBeenRecorded)
						{
							System.Console.WriteLine($"{relativeName} - {holderName} (DUPLICATE)");
							continue;
						}

						System.Console.WriteLine($"{holderName} - {relativeName}");
					}
				}
			}

			// We do a second pass to calculate board member relative links.
			foreach (var companyData in compantDataCollection)
			{
				// Holder name -> Relative name
				HashSet<Tuple<string, string>> relativeLinks = new HashSet<Tuple<string, string>>();
				foreach (var holderNameToRelatives in companyData.RelativeMap)
				{
					string holderName = holderNameToRelatives.Key;
					bool isBoardMember = companyData.BoardAndIndependentDirectors.Contains(holderNameToRelatives.Key);

					foreach (var relativeName in holderNameToRelatives.Value)
					{
						Tuple<string, string> holderNameAndRelativeName = new Tuple<string, string>(holderName, relativeName);
						Tuple<string, string> reversedTuple = new Tuple<string, string>(relativeName, holderName);

						bool hasLinkBeenRecorded = relativeLinks.Contains(reversedTuple) || relativeLinks.Contains(holderNameAndRelativeName);
						if (hasLinkBeenRecorded)
						{
							continue;
						}

						companyData.RelativeLinkCount++;
						relativeLinks.Add(holderNameAndRelativeName);

						bool isRelativeBoardMember = companyData.BoardAndIndependentDirectors.Contains(relativeName);
						if (isBoardMember && isRelativeBoardMember)
						{
							companyData.BoardRelativeLinkCount++;
						}
					}
				}	
			}

			foreach (var companyData in  compantDataCollection)
			{
				System.Console.WriteLine($"{companyData.Name}: " +
					$"\t{companyData.TotalPositionCount} 個職位, " +
					$"\t{companyData.IndependentDirectorCount} 個獨立董事, " +
					$"\t{companyData.BoardPositionCount} 個董事, " +
					$"\t{companyData.RelativeLinkCount} 個親屬連結, " +
					$"\t{companyData.BoardRelativeLinkCount} 個董事親屬連結");
			}

			Console.Read();
		}
	}
}
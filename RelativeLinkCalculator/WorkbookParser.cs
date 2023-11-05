using NanoXLSX;
using RelativeLinkCalculator.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RelativeLinkCalculator
{
	public class CompanyData
	{
		public string Name = "";

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

	public static class WorkbookParser
	{
		const int kCompanyCodeColumn = 0;           // 公司代碼
		const int kDateColumn = 1;                  // 年月日
		const int kStockHolderNameColumn = 2;       // 持股人姓名
		const int kHolderNameColumn = 3;            // 持有人姓名
		const int kPositionColumn = 4;              // 身分別
		const int kRelativeOrSummaryColumn = 5;     // 親屬/總說明

		public static List<CompanyData> ParseCompanyDataWorkbook(Workbook workbook)
		{
			Worksheet sheet = workbook.Worksheets[0];

			IReadOnlyList<Cell> columnA = sheet.GetColumn("A");
			int numberOfEntries = columnA.Count;

			List<CompanyData> companyDataCollection = new List<CompanyData>();
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
						companyDataCollection.Add(currentCompanyData);
					}

					// Reset data counters
					currentCompanyData = new CompanyData();
					currentCompanyData.Name = companyName;

					//System.Console.WriteLine($"-- {companyName} --");
				}

				// Do various counting.

				currentCompanyData.TotalPositionCount++;

				bool hasPositionCell = sheet.TryGetCell(positionAddress, out Cell positionCell) && !positionCell.Value.ToString().IsNullOrEmpty();
				if (hasPositionCell)
				{
					string positionName = positionCell.Value.ToString();

					bool isIndependentDirector = positionName.IsIndependentDirectorPosition();
					if (isIndependentDirector)
					{
						currentCompanyData.IndependentDirectorCount++;
					}

					bool isBoardMember = positionName.IsBoardMemberPosition();
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

				bool isRelativeCell = relativeCell.DataType == Cell.CellType.STRING && relativeCell.Value.ToString().IsCollectionOfRelatives();
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
							//System.Console.WriteLine($"{relativeName} - {holderName} (DUPLICATE)");
							continue;
						}

						//System.Console.WriteLine($"{holderName} - {relativeName}");
					}
				}
			}

			if (!currentCompanyData.Name.IsNullOrEmpty())
			{
				companyDataCollection.Add(currentCompanyData);
			}

			// We do a second pass to calculate board member relative links.
			foreach (var companyData in companyDataCollection)
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

			return companyDataCollection;
		}

		public static void OutputCompanyDataCollectionToWorkbook(string industryCodeStr, List<CompanyData> companyDataCollection, Workbook workbook, string outputPath)
		{
			bool isValidIndustryCode = int.TryParse(industryCodeStr, out int industryCode);
			Worksheet sheet = workbook.CurrentWorksheet;

			Regex companyStockCodePattern = new Regex(@"^\d+");

			int rowIndexOffset = sheet.GetColumn(0).Count;
			for (int rowIndex = 0; rowIndex < companyDataCollection.Count; rowIndex++)
			{
				CompanyData companyData = companyDataCollection[rowIndex];
				int actualRowIndex = rowIndexOffset + rowIndex;

				Match companyStockCodeMatch = companyStockCodePattern.Match(companyData.Name);
				if (!companyStockCodeMatch.Success )
				{
					sheet.AddCell(companyData.Name, 0, actualRowIndex);     // A 
				}
				else
				{
					string stockCodeStr = companyStockCodeMatch.Value;
					bool isValidStockCode = int.TryParse(stockCodeStr, out int stockCode);
					if (isValidStockCode)
					{
						sheet.AddCell(stockCode, 0, actualRowIndex);     // A 
					}
					else
					{
						sheet.AddCell(companyData.Name, 0, actualRowIndex);     // A 
					}
				}

				//sheet.AddCell("(year)", 1, actualRowIndex);					// B

				if (isValidIndustryCode)
				{
					sheet.AddCell(industryCode, 2, actualRowIndex);     // C
				}
				else
				{
					sheet.AddCell(industryCodeStr, 2, actualRowIndex);     // C
				}

				sheet.AddCell(companyData.RelativeLinkCount, 3, actualRowIndex);     // D
				sheet.AddCell(companyData.TotalPositionCount, 4, actualRowIndex);     // E
				sheet.AddCell(companyData.IndependentDirectorCount, 5, actualRowIndex);     // F
				//sheet.AddCell("(ind_board)", 6, actualRowIndex);     // G
				//sheet.AddCell("(general_position)", 7, actualRowIndex);     // H
				//sheet.AddCell("(plinks)", 8, actualRowIndex);     // I
				//sheet.AddCell("(links/plinks)", 9, actualRowIndex);     // J
				sheet.AddCell(companyData.BoardRelativeLinkCount, 10, actualRowIndex);     // K
				sheet.AddCell(companyData.BoardPositionCount, 11, actualRowIndex);     // L
				//sheet.AddCell("(cor_board)", 12, actualRowIndex);     // M
				//sheet.AddCell("(general_board_position)", 13, actualRowIndex);     // N
				//sheet.AddCell("(board_links/gernal_board_position)", 14, actualRowIndex);     // O
			}

			workbook.SaveAs(outputPath);
		}
	}
}

using System;
using System.Text;
using System.Text.RegularExpressions;
using NanoXLSX;
using RelativeLinkCalculator.Extensions;

namespace RelativeLinkCalculator // Note: actual namespace depends on the project name.
{
	internal class Program
	{
		static void Main(string[] args)
		{
			if (args.Length < 3)
			{
				System.Console.WriteLine($"參數必須為 3 個，但是提供的參數只有 {args.Length} 個。");
				System.Console.WriteLine($"參數格式： program.exe <產業代碼> <輸入檔案> <輸出檔案>");
				return;
			}

			System.Console.OutputEncoding = Encoding.Unicode;

			string industryCode = args[0];
			string inputFile = args[1];
			string outputFile = args[2];

			System.Console.WriteLine($"產業： {industryCode}, 輸入: {inputFile}, 輸出: {outputFile}");

			bool isValidInputFile = System.IO.File.Exists(inputFile);
			if (!isValidInputFile) 
			{
				System.Console.WriteLine($"找不到輸入所指定的檔案 '{inputFile}'，程式提早結束。\n");
				return;
			}

			Workbook outputWorkbook = null;
			bool outputFileExists = System.IO.File.Exists(outputFile);
            if (!outputFileExists)
			{
				System.IO.FileInfo outputFileInfo = new FileInfo(outputFile);
				outputFileInfo.Directory.Create();
				
				outputWorkbook = new Workbook(outputFile, "Sheet1");
				outputWorkbook.CurrentWorksheet.AddCell("stock_code", "A1");
				outputWorkbook.CurrentWorksheet.AddCell("year", "B1");
				outputWorkbook.CurrentWorksheet.AddCell("industry", "C1");
				outputWorkbook.CurrentWorksheet.AddCell("general_links", "D1");
				outputWorkbook.CurrentWorksheet.AddCell("total_position", "E1");
				outputWorkbook.CurrentWorksheet.AddCell("ind_board", "F1");
				outputWorkbook.CurrentWorksheet.AddCell("cor_position", "G1");
				outputWorkbook.CurrentWorksheet.AddCell("general_position", "H1");
				outputWorkbook.CurrentWorksheet.AddCell("plinks", "I1");
				outputWorkbook.CurrentWorksheet.AddCell("links/plinks", "J1");
				outputWorkbook.CurrentWorksheet.AddCell("board_links", "K1");
				outputWorkbook.CurrentWorksheet.AddCell("total_board", "L1");
				outputWorkbook.CurrentWorksheet.AddCell("cor_board", "M1");
				outputWorkbook.CurrentWorksheet.AddCell("general_board_position", "N1");
				outputWorkbook.CurrentWorksheet.AddCell("board_links/gernal_board_position", "O1");
				outputWorkbook.CurrentWorksheet.AddCell("TEJ_ family_dummy", "P1");
				outputWorkbook.CurrentWorksheet.AddCell("relation_dummy", "Q1");
				outputWorkbook.Save();
			}
			else
			{
				outputWorkbook = Workbook.Load(outputFile);
			}

            Workbook inputWorkbook = Workbook.Load(inputFile);

			int numberOfEntries = inputWorkbook.CurrentWorksheet.GetColumn("A").Count;
			System.Console.Write($"Number Of Entries: {numberOfEntries}\n");
			var companyDataCollection = WorkbookParser.ParseCompanyDataWorkbook(inputWorkbook);
			foreach (var companyData in companyDataCollection)
			{
				System.Console.WriteLine($"{companyData.Name}: " +
					$"\t{companyData.TotalPositionCount} 個職位, " +
					$"\t{companyData.IndependentDirectorCount} 個獨立董事, " +
					$"\t{companyData.BoardPositionCount} 個董事, " +
					$"\t{companyData.RelativeLinkCount} 個親屬連結, " +
					$"\t{companyData.BoardRelativeLinkCount} 個董事親屬連結");
			}
			System.Console.WriteLine(" ");

			WorkbookParser.OutputCompanyDataCollectionToWorkbook(industryCode, companyDataCollection, outputWorkbook, outputFile);
		}
	}
}
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
			System.Console.OutputEncoding = Encoding.Unicode;
			string fileName = "Input//2022橡膠工業.xlsx";

			Workbook workbook = Workbook.Load(fileName);
			var companyDataCollection = WorkbookParser.ParseCompanyDataWorkbook(workbook);

			foreach (var companyData in companyDataCollection)
			{
				System.Console.WriteLine($"{companyData.Name}: " +
					$"\t{companyData.TotalPositionCount} 個職位, " +
					$"\t{companyData.IndependentDirectorCount} 個獨立董事, " +
					$"\t{companyData.BoardPositionCount} 個董事, " +
					$"\t{companyData.RelativeLinkCount} 個親屬連結, " +
					$"\t{companyData.BoardRelativeLinkCount} 個董事親屬連結");
			}

			System.Console.Read();
		}
	}
}
using NanoXLSX;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RelativeLinkCalculator.Extensions
{
	public static class Extensions
	{
		static readonly Regex kIndependentDirectorPositionPattern = new Regex(@"獨立董事");
		static readonly Regex kBoardPositionPattern = new Regex(@"董事");
		static readonly Regex kGroupHolderColumnPattern = new Regex(@"集團");

		public static bool TryGetCell(this Worksheet sheet, Address address, out Cell cell)
		{
			if (sheet == null) 
			{
				cell = new Cell();
				return false;
			}

			if (sheet.HasCell(address)) 
			{ 
				cell = sheet.GetCell(address);
				return true;
			}

			cell = new Cell();
			return false;
		}

		public static bool IsNullOrEmpty(this string str)
		{
			return str == null || str.Length == 0;
		}

		public static bool IsBoardMemberPosition(this string str)
		{
			return kBoardPositionPattern.IsMatch(str);
		}

		public static bool IsIndependentDirectorPosition(this string str)
		{
			return kIndependentDirectorPositionPattern.IsMatch(str);
		}

		public static bool IsCollectionOfRelatives(this string str)
		{
			return !kGroupHolderColumnPattern.IsMatch(str);
		}

		public static bool TryExtractRelativeName(this string input, out string name)
		{
			// "羅銘鈴(配偶)" -> "羅銘鈴"
			// "羅元隆(四親等)"-> "羅元隆"
			// "陳柏嘉(兄弟姐妹)"-> "陳柏嘉"
			// "陳柏嘉"-> "陳柏嘉"

			string pattern = @"^([^\(\)]+)\s*\(.+?\)?$";

			Match match = Regex.Match(input, pattern);
			if (match.Success)
			{
				name = match.Groups[1].Value.Trim();
				return true;
			}
			else
			{
				// If no match is found, we just return the string with the spaces at the start/end removed.
				name = input.Trim();
				return false;
			}

		}
	}
}

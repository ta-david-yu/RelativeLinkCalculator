using NanoXLSX;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RelativeLinkCalculator.Extensions
{
	internal static class Extensions
	{
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
	}
}

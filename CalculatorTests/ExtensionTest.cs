using NanoXLSX;
using RelativeLinkCalculator;
using RelativeLinkCalculator.Extensions;
using System.Text;

namespace CalculatorTests
{
	[TestClass]
	public class ExtensionTest
	{
		[TestMethod]
		public void IsRelativesOrCollective()
		{
			string group1 = "公開,同一集團(2105)";
			string group2 = "同一集團(2105)";
			string group3 = "上市櫃,同一集團(1440)";
			string group4 = "非同一集團(2389)";
			string group5 = "公開,同一集團(2334)/持股100.00%之子公司";

			string relatives1 = "陳涵馨(配偶),陳秀雄(岳婿)";
			string relatives2 = "陳秀雄(父母),羅銘鈴(父母),陳涵馨(兄弟姐妹),陳柏嘉(兄弟姐妹)";
			string relatives3 = "陳秀雄(一親等),羅元佑(二親等),羅才仁(二親等),陳涵琦(四親等)";
			string relatives4 = "羅才仁(兄弟),陳秀雄(姐夫/妻舅),陳榮華(姐夫/妻舅),羅銘釔(兄妹)";
			string relatives5 = "張宏德(三親等),楊銀明(叔侄),楊信男(叔侄),楊啟仁(叔侄)";
			string relatives6 = "楊啟仁(父女),陳昭榮(四親等),張宏德(四親等),楊信男(叔侄)";

			Assert.IsTrue(!group1.IsCollectionOfRelatives());
			Assert.IsTrue(!group2.IsCollectionOfRelatives());
			Assert.IsTrue(!group3.IsCollectionOfRelatives());
			Assert.IsTrue(!group4.IsCollectionOfRelatives());
			Assert.IsTrue(!group5.IsCollectionOfRelatives());

			Assert.IsTrue(relatives1.IsCollectionOfRelatives());
			Assert.IsTrue(relatives2.IsCollectionOfRelatives());
			Assert.IsTrue(relatives3.IsCollectionOfRelatives());
			Assert.IsTrue(relatives4.IsCollectionOfRelatives());
			Assert.IsTrue(relatives5.IsCollectionOfRelatives());
			Assert.IsTrue(relatives6.IsCollectionOfRelatives());
		}

		[TestMethod]
		public void ExtractRelativeNames1()
		{
			string relatives = "陳秀雄(一親等),羅元佑(二親等),羅才仁(二親等),陳涵琦(四親等)";
			string[] collection = relatives.Split(',');
			Assert.IsTrue(collection.Length == 4);
			for (int i = 0; i < collection.Length; i++)
			{
				bool isValidRelativeName = collection[i].TryExtractRelativeName(out string extractedRelativeName);
				Assert.IsTrue(isValidRelativeName);

				collection[i] = extractedRelativeName;
			}

			Assert.AreEqual("陳秀雄", collection[0]);
			Assert.AreEqual("羅元佑", collection[1]);
			Assert.AreEqual("羅才仁", collection[2]);
			Assert.AreEqual("陳涵琦", collection[3]);
		}

		[TestMethod]
		public void ExtractRelativeNames2()
		{
			string relatives = "羅才仁(兄弟),陳秀雄(姐夫/妻舅),陳榮華(姐夫/妻舅),羅銘釔(兄妹)";
			string[] collection = relatives.Split(',');
			Assert.IsTrue(collection.Length == 4);
			for (int i = 0; i < collection.Length; i++)
			{
				bool isValidRelativeName = collection[i].TryExtractRelativeName(out string extractedRelativeName);
				Assert.IsTrue(isValidRelativeName);

				collection[i] = extractedRelativeName;
			}

			Assert.AreEqual("羅才仁", collection[0]);
			Assert.AreEqual("陳秀雄", collection[1]);
			Assert.AreEqual("陳榮華", collection[2]);
			Assert.AreEqual("羅銘釔", collection[3]);
		}

		[TestMethod]
		public void ParseWorkbook1()
		{
			System.Console.OutputEncoding = Encoding.Unicode;
			string fileName = "TestFiles//2022橡膠工業.xlsx";

			Workbook workbook = Workbook.Load(fileName);
			var companyDataCollection = WorkbookParser.ParseCompanyDataWorkbook(workbook);

			int numberOfCompanies = companyDataCollection.Count;
			Assert.AreEqual(11, numberOfCompanies);

			Assert.AreEqual("2101 南港",		companyDataCollection[0].Name);
			Assert.AreEqual("2102 泰豐",		companyDataCollection[1].Name);
			Assert.AreEqual("2103 台橡",		companyDataCollection[2].Name);
			Assert.AreEqual("2104 國際中橡", companyDataCollection[3].Name);
			Assert.AreEqual("2105 正新",		companyDataCollection[4].Name);
			Assert.AreEqual("2106 建大",		companyDataCollection[5].Name);
			Assert.AreEqual("2107 厚生",		companyDataCollection[6].Name);
			Assert.AreEqual("2108 南帝",		companyDataCollection[7].Name);
			Assert.AreEqual("2109 華豐",		companyDataCollection[8].Name);
			Assert.AreEqual("2114 鑫永銓",	companyDataCollection[9].Name);
			Assert.AreEqual("6582 申豐",		companyDataCollection[10].Name);

			Assert.AreEqual(0, companyDataCollection[0].RelativeLinkCount);
			Assert.AreEqual(0, companyDataCollection[1].RelativeLinkCount);
			Assert.AreEqual(0, companyDataCollection[2].RelativeLinkCount);
			Assert.AreEqual(0, companyDataCollection[3].RelativeLinkCount);
			Assert.AreEqual(31, companyDataCollection[4].RelativeLinkCount);
			Assert.AreEqual(26, companyDataCollection[5].RelativeLinkCount);
			Assert.AreEqual(10, companyDataCollection[6].RelativeLinkCount);
			Assert.AreEqual(7, companyDataCollection[7].RelativeLinkCount);
			Assert.AreEqual(4, companyDataCollection[8].RelativeLinkCount);
			Assert.AreEqual(8, companyDataCollection[9].RelativeLinkCount);
			Assert.AreEqual(0, companyDataCollection[10].RelativeLinkCount);

			Assert.AreEqual(0, companyDataCollection[0].BoardRelativeLinkCount);
			Assert.AreEqual(0, companyDataCollection[1].BoardRelativeLinkCount);
			Assert.AreEqual(0, companyDataCollection[2].BoardRelativeLinkCount);
			Assert.AreEqual(0, companyDataCollection[3].BoardRelativeLinkCount);
			Assert.AreEqual(2, companyDataCollection[4].BoardRelativeLinkCount);
			Assert.AreEqual(11, companyDataCollection[5].BoardRelativeLinkCount);
			Assert.AreEqual(3, companyDataCollection[6].BoardRelativeLinkCount);
			Assert.AreEqual(7, companyDataCollection[7].BoardRelativeLinkCount);
			Assert.AreEqual(0, companyDataCollection[8].BoardRelativeLinkCount);
			Assert.AreEqual(1, companyDataCollection[9].BoardRelativeLinkCount);
			Assert.AreEqual(0, companyDataCollection[10].BoardRelativeLinkCount);
		}
	}
}
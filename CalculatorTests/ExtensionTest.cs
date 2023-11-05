using RelativeLinkCalculator.Extensions;

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

			Assert.AreEqual(collection[0], "陳秀雄");
			Assert.AreEqual(collection[1], "羅元佑");
			Assert.AreEqual(collection[2], "羅才仁");
			Assert.AreEqual(collection[3], "陳涵琦");
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

			Assert.AreEqual(collection[0], "羅才仁");
			Assert.AreEqual(collection[1], "陳秀雄");
			Assert.AreEqual(collection[2], "陳榮華");
			Assert.AreEqual(collection[3], "羅銘釔");
		}
	}
}
namespace CommonLibraryCoreTests;

public class ExcelWrapperReadUnitTest
{
	private IExcelWrapper wrapper;
	[SetUp]
	public void Setup()
	{
	  wrapper = new ExcelWrapper();
	}

	[Test]
	public async Task ReadXlsxFileTest()
	{
	  var path = Path.Combine(Environment.CurrentDirectory, "ExcelFiles");

	  var fn = Path.Combine(path, "bfyyqd.xlsx");
	  var shd = wrapper.ReadExlsSheetDataFromXlsxFile(new ReadXlsxFileParams(fn, 0, 2))?.ToList<ImportTestData>();

	  Assert.NotNull(shd);

	  fn = Path.Combine(path, DateTime.Now.ToString("ddHHmmss") + ".xlsx");
	  var sheetData = shd.ToExlsSheetData(true);
	  await wrapper.WriteToXlsxFileAsync(new WriteXlsxFileParams(sheetData, "fsdf", fn));

	  Assert.IsTrue(File.Exists(fn));
	}
	[Test]
	public void ReadXlsxFileTest2()
	{
	  var path = Path.Combine(Environment.CurrentDirectory, "ExcelFiles");

	  var fn = Path.Combine(path, "importtest.xlsx");
	  var shd = wrapper.ReadExlsSheetDataFromXlsxFile(new ReadXlsxFileParams(fn, 0, 2))?.ToList<ImportTestData>();

	  Assert.NotNull(shd);
	}
}

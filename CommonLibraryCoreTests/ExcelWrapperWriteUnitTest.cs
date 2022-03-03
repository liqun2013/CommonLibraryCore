namespace CommonLibraryCoreTests;

public class ExcelWrapperWriteUnitTest
{
	private IExcelWrapper wrapper;
	[SetUp]
	public void Setup()
	{
	  wrapper = new ExcelWrapper();
	}

	[Test]
	public void WriteToXlsxFileTest()
	{
	  var sheetData1 = new List<ExportTestData>()
	  {
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风2", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风3", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风4", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风5", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风6", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风7", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风8", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风9", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风0", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风10", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风11", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风12", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风13", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 }
	  }.ToExlsSheetData(true);
	  var sheetData2 = new List<ExportTestData>()
	  {
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风2", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风3", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风4", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风5", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风6", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风7", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 23131232321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风8", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风9", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风0", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风10", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 23131342321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风11", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风12", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风13", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 }
	  }.ToExlsSheetData(false);

	  var path = Path.Combine(Environment.CurrentDirectory, "ExcelFiles");
	  if (!Directory.Exists(path))
		Directory.CreateDirectory(path);

	  var fn = Path.Combine(path, DateTime.Now.ToString("ddHHmmss") + ".xlsx");
	  wrapper.WriteToXlsxFile(new WriteXlsxFileParams(new List<ExlsSheetData>
		{
		  sheetData1,sheetData2
		},
		new List<string> { "sheet1", "sheet2" },
		fn,
		false, false)
	  );

	  Assert.IsTrue(File.Exists(fn));
	}
	[Test]
	public void WriteToXlsxFileMergeTest()
	{
	  var sheetData1 = new ExlsSheetData();
	  var cells = new List<SheetCellItem>();
	  cells.Add(new SheetCellItem()
	  {
		ColIndex = 1,
		Data = "col1",
		DataType = DataTypes.SharedString,
		MergeToColIndex = 3,
		RowIndex = 1,
		Texts = new CellTextPart[]
		{
		  new CellTextPart { Text = "col1", TheDataType = DataTypes.String },
		  new CellTextPart { Text = "col3", TheDataType = DataTypes.String }
		}
	  });
	  cells.Add(new SheetCellItem()
	  {
		ColIndex = 2,
		Data = "col2",
		DataType = DataTypes.String,
		RowIndex = 1
	  });
	  cells.Add(new SheetCellItem()
	  {
		ColIndex = 3,
		Data = "col3",
		DataType = DataTypes.String,
		RowIndex = 1
	  });
	  var r = new SheetRowItem(cells, 1);
	  r.RowHeight = 50;
	  sheetData1.AddRow(r);

	  var path = Path.Combine(Environment.CurrentDirectory, "ExcelFiles");
	  if (!Directory.Exists(path))
		Directory.CreateDirectory(path);

	  var fn = Path.Combine(path, DateTime.Now.ToString("ddHHmmss") + ".xlsx");
	  wrapper.WriteToXlsxFile(new WriteXlsxFileParams(new List<ExlsSheetData>
		{
		  sheetData1
		},
		new List<string> { "sheet1" },
		fn,
		true, false)
	  );

	  Assert.IsTrue(File.Exists(fn));
	}
	[Test]
	public void WriteToXlsxFileStyleTest()
	{
	  var sheetData1 = new ExlsSheetData()
	  {
		HeadStyle = new StyleItem
		{
		  BorderStyle = new BorderStyle { BorderColor = "787878", LeftBorderStyle = BorderStyles.Thin, RightBorderStyle = BorderStyles.Medium },
		  FillStyle = new FillStyle { FillColor = "b0b0b0", PatternType = Patterns.Gray125 },
		  FontStyle = new FontStyle { FontBold = true, FontColor = "ff0000", FontName = "宋体", FontSize = 18 },
		  AlignmentStyle = new AlignmentStyle { HorizontalAlignment = HorizontalAlignments.Center, VerticalAlignment = VerticalAlignments.Bottom }
		},
		DataStyle = new StyleItem
		{
		  BorderStyle = new BorderStyle { BorderColor = "787878", LeftBorderStyle = BorderStyles.Thin, RightBorderStyle = BorderStyles.Medium },
		  FillStyle = new FillStyle { PatternType = Patterns.Solid },
		  FontStyle = new FontStyle { FontColor = "ff0000", FontName = "黑体", FontSize = 16 },
		  AlignmentStyle = new AlignmentStyle { WrapText = true, HorizontalAlignment = HorizontalAlignments.Left, VerticalAlignment = VerticalAlignments.Center }
		}
	  };
	  var cells = new List<SheetCellItem>();
	  cells.Add(new SheetCellItem()
	  {
		ColIndex = 1,
		Data = "col1",
		DataType = DataTypes.SharedString,
		MergeToColIndex = 3,
		RowIndex = 1,
		Texts = new CellTextPart[]
		{
		  new CellTextPart { Text = "col1", TheDataType = DataTypes.String },
		  new CellTextPart { Text = "col3", TheDataType = DataTypes.String }
		}
	  });
	  cells.Add(new SheetCellItem()
	  { ColIndex = 2, Data = "col2", DataType = DataTypes.String, RowIndex = 1 });
	  cells.Add(new SheetCellItem()
	  { ColIndex = 3, Data = "col3", DataType = DataTypes.String, RowIndex = 1 });
	  var r = new SheetRowItem(cells, 1);
	  r.RowHeight = 50;
	  sheetData1.AddRow(r);

	  var scells = new List<SheetCellItem>();
	  scells.Add(new SheetCellItem()
	  {
		ColIndex = 1,
		Data = "col1",
		DataType = DataTypes.String,
		RowIndex = 2
	  });
	  scells.Add(new SheetCellItem()
	  { ColIndex = 2, Data = "col2", DataType = DataTypes.String, RowIndex = 2 });
	  scells.Add(new SheetCellItem()
	  { ColIndex = 3, Data = "col3", DataType = DataTypes.String, RowIndex = 2 });
	  var r2 = new SheetRowItem(scells, 2);
	  r2.RowHeight = 50;
	  sheetData1.AddRow(r2);

	  var r3 = new SheetRowItem(new List<SheetCellItem>
	  {
		new SheetCellItem{ ColIndex = 1, Data = "的分为氛围乳房违法分为氛围分为氛围发违反微软否认隔热隔热给微软威风威风", DataType = DataTypes.String, RowIndex = 3,
		 CellStyle = new StyleItem
		 {
		  AlignmentStyle = new AlignmentStyle{ WrapText = false, HorizontalAlignment = HorizontalAlignments.Right, VerticalAlignment = VerticalAlignments.Justify },
		  BorderStyle = new BorderStyle{ BorderColor = "878787", BottomBorderStyle = BorderStyles.Dashed, LeftBorderStyle = BorderStyles.Dashed, RightBorderStyle = BorderStyles.Dotted, TopBorderStyle = BorderStyles.Dotted },
		  FillStyle = new FillStyle{ FillColor = "c0c0c0", PatternType = Patterns.LightGray },
		  FontStyle = new FontStyle{ FontBold = true, FontColor = "101010", FontName = "Microsoft YaHei", FontSize = 30 }
		 }
		},
		new SheetCellItem{ ColIndex = 2, Data = "的分为氛围乳房违法分为氛围分为氛围发违反微软否认隔热隔热给微软威风威风", DataType = DataTypes.String, RowIndex = 3},
		new SheetCellItem{ ColIndex = 3, Data = "的分为氛围乳房违法分为氛围分为氛围发违反微软否认隔热隔热给微软威风威风", DataType = DataTypes.String, RowIndex = 3},
	  }, 3);
	  sheetData1.AddRow(r3);

	  var sheetData2 = new ExlsSheetData()
	  {
		HeadStyle = new StyleItem
		{
		  BorderStyle = new BorderStyle { BorderColor = "0e96fb", LeftBorderStyle = BorderStyles.Thin, RightBorderStyle = BorderStyles.Medium },
		  FillStyle = new FillStyle { FillColor = "0e96fb", PatternType = Patterns.Gray125 },
		  FontStyle = new FontStyle { FontBold = true, FontColor = "000", FontName = "宋体", FontSize = 14 },
		  AlignmentStyle = new AlignmentStyle { HorizontalAlignment = HorizontalAlignments.Center, VerticalAlignment = VerticalAlignments.Bottom }
		},
		DataStyle = new StyleItem
		{
		  BorderStyle = new BorderStyle { BorderColor = "0e96fb", LeftBorderStyle = BorderStyles.Thin, RightBorderStyle = BorderStyles.Medium },
		  FillStyle = new FillStyle { FillColor = "c3e4ff", PatternType = Patterns.Solid },
		  FontStyle = new FontStyle { FontColor = "000", FontName = "黑体", FontSize = 12 },
		  AlignmentStyle = new AlignmentStyle { WrapText = true, HorizontalAlignment = HorizontalAlignments.Left, VerticalAlignment = VerticalAlignments.Center }
		}
	  };
	  _ = sheetData2.ToSheetdata(new List<ExportTestData>()
	  {
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风2", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风3", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风4", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风5", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风6", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风7", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 23131232321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风8", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风9", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风0", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风10", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 23131342321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风11", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风12", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 },
		new ExportTestData{ BoardId = Guid.NewGuid().ToString(), BoxId = Guid.NewGuid().ToString(), Code = "saasasffeewfew瑟夫威风威风13", CreateTime = DateTime.Now, Creator = "me", Description = "啊时代的无法", Height = 12.32m, HourlyUnitFee = 231312321.223m, IsDeleted = true, Length = 231.2333m, Name = "sadsaafre", Status = 1 }
	  }, false);

	  var path = Path.Combine(Environment.CurrentDirectory, "ExcelFiles");
	  if (!Directory.Exists(path))
		Directory.CreateDirectory(path);

	  var fn = Path.Combine(path, DateTime.Now.ToString("ddHHmmss") + ".xlsx");
	  wrapper.WriteToXlsxFile(new WriteXlsxFileParams(new List<ExlsSheetData>
		{
		  sheetData1, sheetData2
		},
		new List<string> { "sheet1", "sheet2" },
		fn,
		true, true)
	  );

	  Assert.IsTrue(File.Exists(fn));
	}

	[Test]
	public async Task TestDataAsync()
	{
	  var lstdata = Enumerable.Range(0, 50000).Select(x => new ImportTestData
	  {
		Ac = $"testdata{x}",
		Code = $"code{x}",
		Cus = $"cus{x}",
		CusMobile = $"1872837483{x}",
		Desc = $"{x}撒大苏打我的去的房产税擦拭东西分为父亲我的确是喜爱程度轻微的测去问问大学生的擦",
		Dl = $"Disdadewdwefwdcsca{x}",
		Dt = DateTime.Now.AddMinutes(x),
		Id = x,
		Mo = $"{x}modfweefvfewewer",
		Mobile = "18347284992",
		Name = $"{x}两岸三地",
		Num = $"{x}sedfewfewfwsdc",
		Remark = "sadwwefwedfsadxczdsvwqefrgbfvdsc",
		Tp = $"{x}sdfwefwefwedcsed",
		Remark2 = "山地车v热v访问测得我威风威风威风成为纷纷威风威风威风为分为氛围分为氛围测"
	  });

	  ExlsSheetData sheetData = lstdata.ToExlsSheetData(true);

	  Assert.AreEqual(sheetData.SheetRows.Count, 50001);

	  var path = Path.Combine(Environment.CurrentDirectory, "ExcelFiles");
	  if (!Directory.Exists(path))
		Directory.CreateDirectory(path);
	  var fn = Path.Combine(path, DateTime.Now.ToString("ddHHmmss") + ".xlsx");
	  await wrapper.WriteToXlsxFileAsync(new WriteXlsxFileParams(sheetData, "sh1", fn));

	  Assert.IsTrue(File.Exists(fn));
	}
	[Test]
	public void TestData()
	{
	  var lstdata = Enumerable.Range(0, 50000).Select(x => new ImportTestData
	  {
		Ac = $"testdata{x}",
		Code = $"code{x}",
		Cus = $"cus{x}",
		CusMobile = $"1872837483{x}",
		Desc = $"{x}撒大苏打我的去的房产税擦拭东西分为父亲我的确是喜爱程度轻微的测去问问大学生的擦",
		Dl = $"Disdadewdwefwdcsca{x}",
		Dt = DateTime.Now.AddMinutes(x),
		Id = x,
		Mo = $"{x}modfweefvfewewer",
		Mobile = "18347284992",
		Name = $"{x}两岸三地",
		Num = $"{x}sedfewfewfwsdc",
		Remark = "sadwwefwedfsadxczdsvwqefrgbfvdsc",
		Tp = $"{x}sdfwefwefwedcsed",
		Remark2 = "山地车v热v访问测得我威风威风威风成为纷纷威风威风威风为分为氛围分为氛围测"
	  });

	  ExlsSheetData sheetData = lstdata.ToExlsSheetData(true);

	  Assert.AreEqual(sheetData.SheetRows.Count, 50001);

	  var path = Path.Combine(Environment.CurrentDirectory, "ExcelFiles");
	  if (!Directory.Exists(path))
		Directory.CreateDirectory(path);
	  var fn = Path.Combine(path, DateTime.Now.ToString("ddHHmmss") + ".xlsx");
	  wrapper.WriteToXlsxFile(new WriteXlsxFileParams(sheetData, "sh1", fn));

	  Assert.IsTrue(File.Exists(fn));
	}
}

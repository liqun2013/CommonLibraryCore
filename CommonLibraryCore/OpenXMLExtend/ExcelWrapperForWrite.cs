namespace CommonLibraryCore;
/// <summary>
/// 这边都是写入到Excel的方法
/// </summary>
public partial class ExcelWrapper : IExcelWrapper
{
  private WriteXlsxFileParams writeParams;
  /// <summary>
  /// 生成多个sheet的Excel文件
  /// </summary>
  public async Task WriteToXlsxFileAsync(WriteXlsxFileParams wxfPrms)
  {
	using var stream = WriteToMemoryStream(wxfPrms);
	using FileStream fs = new(wxfPrms.FileName,
						   FileMode.OpenOrCreate,
						   FileAccess.Write,
						   FileShare.Read,
						   1024 * 2000,
						   options: FileOptions.Asynchronous);
	_ = stream.Seek(0, SeekOrigin.Begin);
	await stream.CopyToAsync(fs).ConfigureAwait(false);
	await fs.FlushAsync().ConfigureAwait(false);
  }

  /// <summary>
  /// 生成多个sheet的Excel文件
  /// </summary>
  public void WriteToXlsxFile(WriteXlsxFileParams wxfPrms)
  {
	using var stream = WriteToMemoryStream(wxfPrms);
	using FileStream fs = new(wxfPrms.FileName,
						   FileMode.OpenOrCreate,
						   FileAccess.Write,
						   FileShare.Read,
						   1024 * 2000,
						   options: FileOptions.Asynchronous);
	long v = stream.Seek(0, SeekOrigin.Begin);
	stream.CopyTo(fs);
	fs.Flush();
  }

  public MemoryStream WriteToMemoryStream(WriteXlsxFileParams wxfPrms)
  {
	writeParams = new(wxfPrms);
	MemoryStream ms = new();
	using SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
	WorkbookPart wbPart = spreadsheetDoc.AddWorkbookPart();
	wbPart.Workbook = new Workbook();
	Sheets sheets = wbPart.Workbook.AppendChild(new Sheets());

	if (writeParams.SheetData?.Any(x => !x.IsEmpty) == true)
	{
	  WorkbookStylesPart stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
	  var sheetindex_formatindex = GenerateSheetStyleForSheets();//将设置的样式转换成openxml格式
	  List<(int sheetindex, uint rowindex, uint colindex, uint formatindex)> cellindex_formatindex = null;
	  if (writeParams.CustFormat)//如果有自定义样式
		cellindex_formatindex = GenerateSheetStyleForCells();

	  SharedStringTablePart sstbpart = writeParams.SheetData?.Any(x => x.IsSharedStringExist) == true ?
									  wbPart.AddNewPart<SharedStringTablePart>() : null;
	  for (uint i = 0; i < writeParams.SheetData.Count(); i++)
	  {
		WorksheetPart worksheetPart = wbPart.AddNewPart<WorksheetPart>();

		if (writeParams.SheetData.ElementAt((int)i) is ExlsSheetData xlsxSheetData)
		{
		  worksheetPart.Worksheet = new Worksheet();
		  if (xlsxSheetData.FirstRow?.RowCells?.Any(x => x.CustWidth > 0) == true)
			GenerateColumns(worksheetPart, xlsxSheetData);

		  //将样式设置到cell
		  if (sheetindex_formatindex?.Any(x => x.sheetindex.Equals((int)i)) == true)
			foreach (var cell in xlsxSheetData.AllCells)
			{
			  var (_, headformatindex, dataformatindex) = sheetindex_formatindex.First(x => x.sheetindex.Equals((int)i));
			  cell.StyleIndex = (xlsxSheetData.FirstRow != null && cell.RowIndex.Equals(xlsxSheetData.FirstRow.RowIndex)) ?
								  headformatindex : dataformatindex;
			}

		  //将自定义样式设置到cell
		  if (writeParams.CustFormat && cellindex_formatindex?.Any() == true)
			foreach (var cell in xlsxSheetData.AllCells)
			  if (cellindex_formatindex.Any(x => x.sheetindex.Equals((int)i) && x.rowindex.Equals(cell.RowIndex) && x.colindex.Equals(cell.ColIndex)))
				cell.StyleIndex = cellindex_formatindex.First(x => x.sheetindex.Equals((int)i) && x.rowindex.Equals(cell.RowIndex) && x.colindex.Equals(cell.ColIndex))
														 .formatindex;

		  worksheetPart.Worksheet.Append(CreateSheetData(xlsxSheetData, sstbpart));
		  if (writeParams.DoMerge)
			GenerateMergeCells(worksheetPart, xlsxSheetData);
		}

		_ = sheets.AppendChild(new Sheet
		{ Id = wbPart.GetIdOfPart(worksheetPart), SheetId = i + 1, Name = writeParams.SheetNames?.ElementAt((int)i) ?? string.Empty });
	  }

	  stylesPart.Stylesheet = new Stylesheet(fonts, fills, borders, cellFormats);
	}
	wbPart.Workbook.Save();
	return ms;
  }

  /// <summary>
  /// 有设置列宽的需要创建列，一般设置在第一行
  /// </summary>
  private void GenerateColumns(WorksheetPart worksheetPart, ExlsSheetData xlsxSheetData)
  {
	Columns cols = worksheetPart.Worksheet.AppendChild(new Columns());
	foreach (var itm in xlsxSheetData.FirstRow.RowCells.Where(x => x.CustWidth > 0).OrderBy(x => x.ColIndex))
	  cols.Append(new Column { CustomWidth = true, Min = itm.ColIndex, Max = itm.ColIndex, Width = itm.CustWidth });
  }

  /// <summary>
  /// 生成一个sheet数据
  /// </summary>
  /// <param name="rowItems"></param>
  /// <returns></returns>
  private SheetData CreateSheetData(ExlsSheetData xlsxSheetData, SharedStringTablePart sharedStringPart)
  {
	SheetData sheetData = (xlsxSheetData.SheetRows?.Any() == true) ? new SheetData() : null;

	if (sheetData != null)
	  foreach (var r in xlsxSheetData.SheetRows.OrderBy(x => x.RowIndex))
		sheetData.AppendChild(CreateSheetRow(r, sharedStringPart));

	return sheetData;
  }
  protected Row CreateSheetRow(SheetRowItem item, SharedStringTablePart shareStringPart)
  {
	Row result = new() { RowIndex = item.RowIndex };
	if (item.RowCells?.Any() == true)
	{
	  if (item.RowHeight > 0)
	  {
		result.CustomHeight = BooleanValue.FromBoolean(true);
		result.Height = item.RowHeight;
	  }

	  foreach (var itm in item.RowCells.OrderBy(x => x.ColIndex))
		_ = result.AppendChild(CreateCell(itm, item.RowIndex, shareStringPart));
	}

	return result;
  }

  private Cell CreateCell(SheetCellItem itm, uint row, SharedStringTablePart ssPart)
  {
	Cell result = null;
	string col = ColumnLetter(itm.ColIndex);
	switch (itm.DataType)
	{
	  case DataTypes.Number:
		result = new() { DataType = CellValues.Number, CellReference = col + row.ToString(), CellValue = new CellValue(itm.Data) };
		break;
	  case DataTypes.String:
		result = new() { DataType = CellValues.InlineString, CellReference = col + row.ToString() };

		InlineString istring = new();
		_ = istring.AppendChild(new Text { Text = itm.Data ?? string.Empty });
		_ = result.AppendChild(istring);
		break;
	  case DataTypes.SharedString:
		var index = GenerateSharedStringItem(itm.Texts, ssPart);
		result = new() { DataType = CellValues.SharedString, CellReference = col + row.ToString(), CellValue = new CellValue(index.ToString()) };
		break;
	}

	if (result != null && itm.StyleIndex > 0)
	  result.StyleIndex = itm.StyleIndex;

	return result;
  }
  /// <summary>
  /// 合并单元格
  /// 注意: 合并单元格只需要设置开始单元格的MergeToRowIndex(合并至单元格的行),MergeToColIndex(合并至单元格的列)属性，
  /// 不需要设置结束单元格, 所以MergeToRowIndex必须大于RowIndex、MergeToColIndex必须大于ColIndex
  /// </summary>
  /// <param name="cellItems"></param>
  /// <param name="worksheet"></param>
  private void GenerateMergeCells(WorksheetPart worksheetPart, ExlsSheetData xlsxSheetData)
  {
	var cellsToMerge = xlsxSheetData.AllCells.Where(x => x.MergeToColIndex > 0 || x.MergeToRowIndex > 0);
	if (cellsToMerge?.Any() == true)
	{
	  MergeCells mergeCells = worksheetPart.Worksheet.AppendChild(new MergeCells());
	  foreach (var itm in cellsToMerge)
	  {
		var mergeFrom = ColumnLetter(itm.ColIndex) //需要将列索引转成Excel的列头，如:1->A,27->A1
						+ itm.RowIndex.ToString();
		var mergeTo = ColumnLetter(Math.Max(itm.ColIndex, itm.MergeToColIndex))  //需要将列索引转成Excel的列头，如:1->A,27->A1
						+ Math.Max(itm.RowIndex, itm.MergeToRowIndex).ToString();
		MergeCell mergeCell = new() { Reference = new StringValue(mergeFrom + ":" + mergeTo) };
		mergeCells.Append(mergeCell);
	  }

	  //worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());
	}
  }
  private int GenerateSharedStringItem(CellTextPart[] textParts, SharedStringTablePart shareStringPart)
  {
	if (shareStringPart.SharedStringTable == null)
	  shareStringPart.SharedStringTable = new SharedStringTable();

	SharedStringItem ssItm = new();

	foreach (var part in textParts)
	{
	  Run run = new();
	  var txt = part.Text;
	  if (part.PartFontStyle != null)
	  {
		RunProperties runProperties = new();
		var scf = part.PartFontStyle;
		if (scf.FontBold)
		  runProperties.Append(new Bold());
		if (scf.FontSize > 0)
		  runProperties.Append(new FontSize { Val = scf.FontSize });
		if (!string.IsNullOrEmpty(scf.FontColor))
		  runProperties.Append(new Color { Rgb = scf.FontColor });
		if (part.TheDataType == DataTypes.Number)
		  runProperties.Append(new NumberingFormat());

		run.Append(runProperties);
	  }
	  Text text = new() { Text = txt };
	  run.Append(text);

	  ssItm.Append(run);
	}
	// The text does not exist in the part. Create the SharedStringItem and return its index.
	shareStringPart.SharedStringTable.AppendChild(ssItm);
	shareStringPart.SharedStringTable.Save();

	// Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
	return shareStringPart.SharedStringTable.ChildElements.Count - 1;
  }

  #region 样式
  private List<(uint index, FontStyle fontstyle)> index_fontstyle;
  private List<(uint index, FillStyle fillstyle)> index_fillstyle;
  private List<(uint index, BorderStyle borderstyle)> index_borderstyle;
  private List<(uint index, AlignmentStyle alignstyle)> index_alignstyle;
  private List<(uint index, StyleItem styleitem)> index_styleitem;

  private Fonts fonts = new(new Font(new FontStyle { FontSize = 10 }.ToOpenXmlFont()));
  private Fills fills = new(new Fill(new FillStyle().ToOpenXmlFill()));
  private Borders borders = new(new BorderStyle().ToOpenXmlBorder());
  private CellFormats cellFormats = new(new CellFormat() { FillId = 0, FontId = 0, BorderId = 0 });

  private uint fontstyleindex = 1;
  private uint fillstyleindex = 1;
  private uint borderstyleindex = 1;
  private uint alignstyleindex = 1;
  private uint styleitmformatindex = 1;
  /// <summary>
  /// 生成sheet统一样式,一个sheet只定义一个表头单元格样式和一个数据单元格样式
  /// </summary>
  /// <param name="stylesPart"></param>
  /// <returns></returns>
  private List<(int sheetindex, uint headformatindex, uint dataformatindex)> GenerateSheetStyleForSheets()
  {
	List<(int sheetindex, uint headformatindex, uint dataformatindex)> result = null;
	if (writeParams.SheetData?.Any(x => x.HeadStyle != null || x.DataStyle != null) == true)
	{
	  result = new();

	  index_fontstyle = new();
	  index_fillstyle = new();
	  index_borderstyle = new();
	  index_alignstyle = new();
	  index_styleitem = new();
	  for (int i = 0; i < writeParams.SheetData.Count(); i++)
	  {
		if (writeParams.SheetData.ElementAt(i) is ExlsSheetData sheet)
		{
		  (int sheetindex, uint headformatindex, uint dataformatindex) ri = new(i, 0, 0);
		  if (sheet.HeadStyle != null)
			ri.headformatindex = GenerateSheetStyleItem(sheet.HeadStyle);
		  if (sheet.DataStyle != null)
			ri.dataformatindex = GenerateSheetStyleItem(sheet.DataStyle);

		  result.Add(ri);
		}
	  }
	}
	return result;
  }

  private uint GenerateSheetStyleItem(StyleItem s)
  {
	var tag = false;
	if (s.HasFontStyle && !index_fontstyle.Any(x => x.fontstyle.Equals(s.FontStyle)))
	{
	  fonts.AppendChild(s.FontStyle.ToOpenXmlFont());
	  index_fontstyle.Add((fontstyleindex++, s.FontStyle));
	  tag = true;
	}
	if (s.HasFillStyle && !index_fillstyle.Any(x => x.fillstyle.Equals(s.FillStyle)))
	{
	  fills.AppendChild(s.FillStyle.ToOpenXmlFill());
	  index_fillstyle.Add((fillstyleindex++, s.FillStyle));
	  tag = true;
	}
	if (s.HasBorderStyle && !index_borderstyle.Any(x => x.borderstyle.Equals(s.BorderStyle)))
	{
	  borders.AppendChild(s.BorderStyle.ToOpenXmlBorder());
	  index_borderstyle.Add((borderstyleindex++, s.BorderStyle));
	  tag = true;
	}
	if (s.HasAlignmentStyle && !index_alignstyle.Any(x => x.alignstyle.Equals(s.AlignmentStyle)))
	{
	  index_alignstyle.Add((alignstyleindex++, s.AlignmentStyle));
	  tag = true;
	}
	if (tag)
	{
	  CellFormat cf = new();
	  if (s.HasFontStyle)
	  {
		cf.ApplyFont = true;
		cf.FontId = index_fontstyle.FirstOrDefault(x => x.fontstyle.Equals(s.FontStyle)).index;
	  }
	  if (s.HasFillStyle)
	  {
		cf.ApplyFill = true;
		cf.FillId = index_fillstyle.FirstOrDefault(x => x.fillstyle.Equals(s.FillStyle)).index;
	  }
	  if (s.HasBorderStyle)
	  {
		cf.ApplyBorder = true;
		cf.BorderId = index_borderstyle.FirstOrDefault(x => x.borderstyle.Equals(s.BorderStyle)).index;
	  }
	  if (s.HasAlignmentStyle)
	  {
		cf.ApplyAlignment = true;
		cf.Append(index_alignstyle.First(x => x.alignstyle.Equals(s.AlignmentStyle)).alignstyle.ToOpenXmlAlignment());
	  }

	  _ = cellFormats.AppendChild(cf);
	  index_styleitem.Add((styleitmformatindex, s));
	  return styleitmformatindex++;
	}
	else if (index_styleitem.Any(x => x.styleitem.Equals(s)))
	  return index_styleitem.First(x => x.styleitem.Equals(s)).index;
	else
	  throw new Exception("unknow styleitem");
  }

  /// <summary>
  /// 生成各个单元格的样式，样式设置在单元格里面的
  /// </summary>
  /// <param name="workbookPart"></param>
  /// <returns></returns>
  private List<(int sheetindex, uint rowindex, uint colindex, uint formatindex)> GenerateSheetStyleForCells()
  {
	List<(int sheetindex, uint rowindex, uint colindex, uint formatindex)> result = null;
	if (writeParams.SheetData?.SelectMany(x => x.AllCells)?.Any(x => x.CellStyle != null) == true)
	{
	  result = new();

	  index_fontstyle = new();
	  index_fillstyle = new();
	  index_borderstyle = new();
	  index_alignstyle = new();
	  index_styleitem = new();
	  for (int i = 0; i < writeParams.SheetData.Count(); i++)
	  {
		if (writeParams.SheetData.ElementAt(i) is ExlsSheetData sheet)
		{
		  for (var ri = 1; ri <= sheet.SheetRows.Count; ri++)
		  {
			if (sheet.SheetRows[ri - 1]?.RowCells.Any(x => x.CellStyle != null) == true)
			{
			  for (var ci = 1; ci <= sheet.SheetRows[ri - 1].RowCells.Count; ci++)
			  {
				var theCellStyle = sheet.SheetRows[ri - 1].RowCells[ci - 1].CellStyle;
				if (theCellStyle != null)
				{
				  var fi = GenerateSheetStyleItem(theCellStyle);
				  result.Add(new(i, (uint)ri, (uint)ci, fi));
				}
			  }
			}
		  }
		}
	  }
	}
	return result;
  }
  #endregion
  /// <summary>
  /// 因为Excel里面的列是字母表示的，所以要把数字列的索引转成Excel里面的列头
  /// 如: 1->A,27->AA
  /// </summary>
  /// <param name="uintCol"></param>
  /// <returns></returns>
  private string ColumnLetter(uint uintCol)
  {
	var intCol = (int)uintCol;
	int intFirstLetter = ((intCol - 1) / 676) + 64;
	int intSecondLetter = (((intCol - 1) % 676) / 26) + 64;
	int intThirdLetter = ((intCol - 1) % 26) + 65;

	char firstLetter = (intFirstLetter > 64) ? (char)intFirstLetter : char.MinValue;//' ';
	char secondLetter = (intSecondLetter > 64) ? (char)intSecondLetter : char.MinValue;//' ';
	char thirdLetter = (char)intThirdLetter;

	var result = string.Concat(firstLetter, secondLetter, thirdLetter).Trim(char.MinValue);
	return result;
  }
}

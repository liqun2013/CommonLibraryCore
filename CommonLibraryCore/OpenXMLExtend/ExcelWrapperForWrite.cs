namespace CommonLibraryCore
{
  public partial class ExcelWrapper : IExcelWrapper
  {
	/// <summary>
	/// 生成多个sheet的Excel文件
	/// </summary>
	public async void WriteToXlsxFileAsync(WriteXlsxFileParams wxfPrms)
	{
	  using (var stream = WriteToMemoryStream(wxfPrms))
	  using (FileStream fs = new(wxfPrms.FileName, FileMode.OpenOrCreate, FileAccess.Write))
	  {
		long v = stream.Seek(0, SeekOrigin.Begin);
		await stream.CopyToAsync(fs).ConfigureAwait(false);
		await fs.FlushAsync().ConfigureAwait(false);
	  }
	}

	/// <summary>
	/// 生成多个sheet的Excel文件
	/// </summary>
	public void WriteToXlsxFile(WriteXlsxFileParams wxfPrms)
	{
	  Task.Factory.StartNew((obj) =>
	  {
		WriteToXlsxFileAsync(obj as WriteXlsxFileParams);
	  }, wxfPrms).ConfigureAwait(false);
	}

	public MemoryStream WriteToMemoryStream(WriteXlsxFileParams wxfPrms)
	{
	  MemoryStream ms = new();
	  using (SpreadsheetDocument? spreadsheetDoc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
	  {
		WorkbookPart workbookPart = spreadsheetDoc.AddWorkbookPart();
		workbookPart.Workbook = new Workbook();
		Sheets sheets = spreadsheetDoc.WorkbookPart?.Workbook.AppendChild(new Sheets());

		var stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
		List<SheetCellItem> allCells = new();
		for (uint i = 0; i < wxfPrms.LstSheetData.Count; i++)
		{
		  if (wxfPrms.LstSheetData[(int)i] is ExlsSheetData xlsxSheetData)
			allCells.AddRange(xlsxSheetData.AllCells);
		}
		if (allCells.Any())
		{
		  stylePart.Stylesheet = (wxfPrms.CustFormat ? GenerateStylesheet(allCells) : DefaultStylesheet())
					?? DefaultStylesheet();
		  stylePart.Stylesheet.Save();
		}

		SharedStringTablePart tbpart = (allCells.Any(x => x.DataType == DataTypes.SharedString)) ?
		  tbpart = workbookPart.AddNewPart<SharedStringTablePart>() : null;
		for (uint i = 0; i < wxfPrms.LstSheetData.Count; i++)
		{
		  WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

		  if (wxfPrms.LstSheetData[(int)i] is ExlsSheetData xlsxSheetData)
		  {
			worksheetPart.Worksheet = new Worksheet();
			if (wxfPrms.CustFormat)
			{
			  var cols = GenerateColumns(xlsxSheetData.FirstRow);
			  if (cols != null)
				worksheetPart.Worksheet.Append(cols);
			}
			else
			{
			  foreach (var c in xlsxSheetData.AllCells)
				c.FormatIndex = c.RowIndex.Equals(1) ? 1u : 2u;
			}

			worksheetPart.Worksheet.Append(CreateSheetData(xlsxSheetData.SheetRows, tbpart));
			if (wxfPrms.DoMerge)
			  DoMerge(worksheetPart.Worksheet, xlsxSheetData.AllCells);
		  }

		  sheets?.AppendChild(new Sheet
		  {
			Id = spreadsheetDoc.WorkbookPart?.GetIdOfPart(worksheetPart),
			SheetId = UInt32Value.FromUInt32(i + 1),
			Name = wxfPrms.LstSheetName[(int)i][..Math.Min(wxfPrms.LstSheetName[(int)i].Length, 30)]
		  });
		}
		workbookPart.Workbook.Save();
	  }
	  return ms;
	}

	private Cell CreateCell(string col, uint row, string txt, DataTypes dType, CellTextPart[] txtPrts, SharedStringTablePart ssPart, uint fmtIndex = 0)
	{
	  Cell result = null;
	  switch (dType)
	  {
		case DataTypes.Number:
		  result = new() { DataType = CellValues.Number, CellReference = col + row.ToString(), CellValue = new CellValue(txt) };
		  break;
		case DataTypes.String:
		  result = new() { DataType = CellValues.InlineString, CellReference = col + row.ToString() };

		  InlineString istring = new();
		  istring.AppendChild(new Text { Text = txt ?? string.Empty });
		  result.AppendChild(istring);
		  break;
		case DataTypes.SharedString:
		  var index = GenerateSharedStringItem(txtPrts, ssPart);
		  result = new() { DataType = CellValues.SharedString, CellReference = col + row.ToString(), CellValue = new CellValue(index.ToString()) };
		  break;
	  }

	  if (result != null && fmtIndex > 0)
		result.StyleIndex = fmtIndex;

	  return result;
	}

	/// <summary>
	/// 因为Excel里面的列是字母表示的，所以要把数字列的索引转成Excel里面的列头
	/// 如: 1->A,27->AA
	/// </summary>
	/// <param name="intCol"></param>
	/// <returns></returns>
	private string ColumnLetter(int intCol)
	{
	  int intFirstLetter = ((intCol - 1) / 676) + 64;
	  int intSecondLetter = (((intCol - 1) % 676) / 26) + 64;
	  int intThirdLetter = ((intCol - 1) % 26) + 65;

	  char firstLetter = (intFirstLetter > 64) ? (char)intFirstLetter : char.MinValue;//' ';
	  char secondLetter = (intSecondLetter > 64) ? (char)intSecondLetter : char.MinValue;//' ';
	  char thirdLetter = (char)intThirdLetter;

	  return string.Concat(firstLetter, secondLetter, thirdLetter)/*.Trim()*/;
	}

	/// <summary>
	/// 生成一个sheet数据
	/// </summary>
	/// <param name="rowItems"></param>
	/// <returns></returns>
	private SheetData CreateSheetData(IEnumerable<SheetRowItem> rowItems, SharedStringTablePart sharedStringPart)
	{
	  SheetData sheetData = (rowItems?.Any() == true) ? new SheetData() : null;

	  if (sheetData != null)
		foreach (var r in rowItems.OrderBy(x => x.RowIndex))
		  sheetData.AppendChild(CreateSheetRow(r, sharedStringPart));

	  return sheetData;
	}
	/// <summary>
	/// 合并单元格
	/// 注意: 合并单元格只需要设置开始单元格的MergeToRowIndex(合并至单元格的行),MergeToColIndex(合并至单元格的列)属性，
	/// 不需要设置结束单元格, 所以MergeToRowIndex必须大于RowIndex、MergeToColIndex必须大于ColIndex
	/// </summary>
	/// <param name="cellItems"></param>
	/// <param name="worksheet"></param>
	private void DoMerge(Worksheet worksheet, IEnumerable<SheetCellItem> cellItems)
	{
	  var cellsToMerge = cellItems.Where(x => x.MergeToColIndex > 0 || x.MergeToRowIndex > 0);
	  if (cellsToMerge?.Any() == true)
	  {
		MergeCells mergeCells = new();
		foreach (var itm in cellsToMerge)
		{
		  var mergeFrom = ColumnLetter((int)itm.ColIndex) //需要将列索引转成Excel的列头，如:1->A,27->A1
						  + itm.RowIndex.ToString();
		  var mergeTo = ColumnLetter((int)Math.Max(itm.ColIndex, itm.MergeToColIndex))  //需要将列索引转成Excel的列头，如:1->A,27->A1
						  + Math.Max(itm.RowIndex, itm.MergeToRowIndex).ToString();
		  MergeCell mergeCell = new() { Reference = new StringValue(mergeFrom + ":" + mergeTo) };
		  mergeCells.Append(mergeCell);
		}

		worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
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
		if (part.PartFormat != null)
		{
		  RunProperties runProperties = new();
		  var scf = part.PartFormat;
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
	private int FindFont(Fonts fonts, Font font)
	{
	  for (int i = 0; i < fonts.ChildElements.Count; i++)
	  {
		if (fonts.ChildElements[i].OuterXml.Equals(font.OuterXml, StringComparison.OrdinalIgnoreCase))
		  return i;
	  }
	  return -1;
	}
	private int FindForGroundFill(Fills fills, Fill fill)
	{
	  for (int i = 0; i < fills.ChildElements.Count; i++)
	  {
		if (fills.ChildElements[i].OuterXml.Equals(fill.OuterXml, StringComparison.OrdinalIgnoreCase))
		  return i;
	  }
	  return -1;
	}

	private Stylesheet GenerateStylesheet(IEnumerable<SheetCellItem> cellItems)
	{
	  if (cellItems?.Any(x => x.CellFormats != null) == true)
	  {
		IEnumerable<SheetCellFormats> cfs = cellItems.Where(x => x.CellFormats != null).Select(x => x.CellFormats).Distinct();
		Fills fills = new(new Fill(new PatternFill() { PatternType = PatternValues.None }),
							new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }));
		Dictionary<string, Fill> fgColorAndFills = new();
		if (cfs.Any(x => !string.IsNullOrEmpty(x.FGColor)))
		{
		  foreach (var fg in cfs.Where(x => !string.IsNullOrEmpty(x.FGColor)).Select(x => x.FGColor).Distinct())
		  {
			var f = new Fill(new PatternFill() { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor() { Rgb = new HexBinaryValue(fg) } });
			fgColorAndFills.Add(fg, f);
			fills.AppendChild(f);
		  }
		}
		Fonts fonts = new(new Font(new FontSize() { Val = 10 }),
														new Font(new FontSize() { Val = 12 }, new Bold() { Val = BooleanValue.FromBoolean(true) }),
														new Font(new FontSize() { Val = 10 }, new FontName() { Val = new StringValue("Arial") }));
		Borders borders = new(new Border(),
			new Border(
							new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
							new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
							new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
							new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
							new DiagonalBorder()),
			new Border(new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin }),
			new Border(new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin }),
			new Border(new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin }),
			new Border(new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin }));
		CellFormat[] lstCellFormats = new CellFormat[cfs.Count() + 1];
		lstCellFormats[0] = (new CellFormat() { FillId = 0, FontId = 0 });
		uint index = 1;
		bool tag = false;
		foreach (SheetCellFormats itm in cfs)
		{
		  tag = false;
		  CellFormat cf = new();
		  if (itm.FontBold || itm.FontSize > 0 || !string.IsNullOrEmpty(itm.FontColor) || !string.IsNullOrEmpty(itm.FontName))
		  {
			var f = new Font();
			if (itm.FontSize > 0)
			  f.Append(new FontSize() { Val = itm.FontSize });
			if (itm.FontBold)
			  f.Append(new Bold() { Val = true });
			if (!string.IsNullOrEmpty(itm.FontColor))
			  f.Append(new Color { Rgb = itm.FontColor });
			if (!string.IsNullOrEmpty(itm.FontName))
			  f.Append(new FontName { Val = itm.FontName });

			int i = FindFont(fonts, f);
			if (i > -1)
			{
			  cf.FontId = (uint)i;
			}
			else
			{
			  fonts.Append(f);
			  cf.FontId = (uint)(fonts.ChildElements.Count - 1);
			}
			cf.ApplyFont = true;
			tag = true;
		  }

		  if (itm.HorizontalAlignment != HorizontalAlignments.Default || itm.VerticalAlignment != VerticalAlignments.Default || itm.WrapText)
		  {
			var align = new Alignment();
			switch (itm.HorizontalAlignment)
			{
			  case HorizontalAlignments.Default:
				break;
			  case HorizontalAlignments.Center:
				align.Horizontal = HorizontalAlignmentValues.Center;
				break;
			  case HorizontalAlignments.Left:
				align.Horizontal = HorizontalAlignmentValues.Left;
				break;
			  case HorizontalAlignments.Right:
				align.Horizontal = HorizontalAlignmentValues.Right;
				break;
			}
			switch (itm.VerticalAlignment)
			{
			  case VerticalAlignments.Default:
				break;
			  case VerticalAlignments.Top:
				align.Vertical = VerticalAlignmentValues.Top;
				break;
			  case VerticalAlignments.Middle:
				align.Vertical = VerticalAlignmentValues.Center;
				break;
			  case VerticalAlignments.Bottom:
				align.Vertical = VerticalAlignmentValues.Bottom;
				break;
			}
			align.WrapText = itm.WrapText;
			cf.Append(align);
			cf.ApplyAlignment = true;
			tag = true;
		  }

		  if (tag)
		  {
			cf.FillId = 0;
			lstCellFormats[index] = cf;
		  }
		  index++;
		}

		index = 1;
		foreach (SheetCellFormats itm in cfs)
		{
		  if (itm.Borders != null && itm.Borders.Any(x => x == true))
		  {
			if (itm.FontBold || itm.FontSize > 0 || !string.IsNullOrEmpty(itm.FontColor) || !string.IsNullOrEmpty(itm.FontName) ||
				itm.HorizontalAlignment != HorizontalAlignments.Default || itm.VerticalAlignment != VerticalAlignments.Default)
			{
			  if (itm.Borders[0] && itm.Borders[1] && itm.Borders[2] && itm.Borders[3])
				lstCellFormats[index].BorderId = 1;
			  else if (itm.Borders[0])
				lstCellFormats[index].BorderId = 2;
			  else if (itm.Borders[1])
				lstCellFormats[index].BorderId = 3;
			  else if (itm.Borders[2])
				lstCellFormats[index].BorderId = 4;
			  else if (itm.Borders[3])
				lstCellFormats[index].BorderId = 5;
			}
			else
			{
			  CellFormat fc = new();
			  if (itm.Borders[0] && itm.Borders[1] && itm.Borders[2] && itm.Borders[3])
				fc.BorderId = 1;
			  else if (itm.Borders[0])
				fc.BorderId = 2;
			  else if (itm.Borders[1])
				fc.BorderId = 3;
			  else if (itm.Borders[2])
				fc.BorderId = 4;
			  else if (itm.Borders[3])
				fc.BorderId = 5;
			  lstCellFormats[index] = fc;
			}
			lstCellFormats[index].ApplyBorder = true;
		  }
		  if (!string.IsNullOrEmpty(itm.FGColor))
			lstCellFormats[index].FillId = (uint)FindForGroundFill(fills, fgColorAndFills[itm.FGColor]);
		  index++;
		}
		index = 1;
		foreach (SheetCellFormats itm in cfs)
		{
		  foreach (SheetCellItem x in cellItems.Where(x => x.CellFormats == itm))
			x.FormatIndex = index;

		  index++;
		}
		CellFormats cellFormats = new(lstCellFormats);
		return new Stylesheet(fonts, fills, borders, cellFormats);
	  }
	  return null;
	}

	/// <summary>
	/// 默认样式
	/// </summary>
	/// <param name="sheetDatas"></param>
	/// <returns></returns>
	private Stylesheet DefaultStylesheet()
	{
	  Fonts fonts = new(new Font(new FontSize() { Val = 10 }, new FontName { Val = "宋体" }),
		  new Font(new FontSize() { Val = 10 }, new Bold() { Val = true }, new Color { Rgb = "993366" }, new FontName { Val = "宋体" }));
	  Fills fills = new(new Fill(new PatternFill() { PatternType = PatternValues.None }),
		  new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }),
		  new Fill(new PatternFill() { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor() { Rgb = new HexBinaryValue("C0C0C0") } }));
	  Borders borders = new(new Border(),
		  new Border(new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
						  new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
						  new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
						  new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
						  new DiagonalBorder()));
	  CellFormat headFormat = new() { FontId = 1, ApplyFont = true, FillId = 2, BorderId = 1, ApplyBorder = true };
	  CellFormat dataFormat = new() { FontId = 0, ApplyFont = true, FillId = 0, BorderId = 1, ApplyBorder = true };

	  CellFormats cellFormats = new(new CellFormat() { FillId = 0, FontId = 0 }, headFormat, dataFormat);
	  return new Stylesheet(fonts, fills, borders, cellFormats);
	}

	#endregion

	/// <summary>
	/// 有设置列宽的需要创建列，一般设置在第一行
	/// </summary>
	/// <param name="firstRow"></param>
	/// <returns></returns>
	private Columns GenerateColumns(SheetRowItem firstRow)
	{
	  Columns result = null;
	  if (firstRow?.RowCells?.Any(x => x.CustWidth > 0) == true)
	  {
		result = new Columns();
		foreach (var itm in firstRow.RowCells.Where(x => x.CustWidth > 0).OrderBy(x => x.ColIndex))
		  result.Append(new Column { CustomWidth = true, Min = itm.ColIndex, Max = itm.ColIndex, Width = itm.CustWidth });
	  }
	  return result;
	}
	protected Row CreateSheetRow(SheetRowItem item, SharedStringTablePart shareStringPart)
	{
	  Row result = new Row { RowIndex = item.RowIndex };
	  if (item.RowCells?.Any() == true)
	  {
		if (item.RowHeight > 0)
		{
		  result.CustomHeight = BooleanValue.FromBoolean(true);
		  result.Height = item.RowHeight;
		}

		foreach (var itm in item.RowCells.OrderBy(x => x.ColIndex))
		{
		  Cell cell = CreateCell(ColumnLetter((int)itm.ColIndex), item.RowIndex, itm.Data, itm.DataType, itm.Texts, shareStringPart, itm.FormatIndex);
		  result.AppendChild(cell);
		}
	  }

	  return result;
	}
  }
}

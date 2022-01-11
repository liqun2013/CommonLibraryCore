namespace CommonLibraryCore
{
  public partial class ExcelWrapper : IExcelWrapper
  {
	#region 写入到Excel相关

	/// <summary>
	/// 生成多个sheet的Excel文件
	/// </summary>
	/// <param name="dataItems"></param>
	/// <param name="shNames">sheet name</param>
	/// <param name="filename">生成的excel文件名(包含路径)</param>
	/// <param name="doMerge">true: 存在合并单元格/false: 不存在合并单元格</param>
	/// <param name="customFormat">true: 有自定义样式/false: 没有自定义样式</param>
	public void GenerateXlsxFile(List<BaseSheetData> exlsSheetData, string[] shNames, string filename, bool doMerge = false, bool customFormat = false)
	{
	  using SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook);
	  WorkbookPart workbookPart = spreadsheetDoc.AddWorkbookPart();
	  workbookPart.Workbook = new Workbook();
	  Sheets sheets = spreadsheetDoc.WorkbookPart?.Workbook.AppendChild(new Sheets());
	  SharedStringTablePart tbpart = null;

	  var stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
	  List<SheetCellItem> allCells = new();
	  for (uint i = 0; i < exlsSheetData.Count; i++)
	  {
		if (exlsSheetData[(int)i] is ExlsSheetData xlsxSheetData)
		  allCells.AddRange(xlsxSheetData.AllCells);
	  }
	  if (allCells.Any())
	  {
		if (customFormat)
		  stylePart.Stylesheet = GenerateStylesheet(allCells);
		else
		  stylePart.Stylesheet = DefaultStylesheet();//GenerateDefaultStylesheet(allCells);
		stylePart.Stylesheet.Save();
	  }

	  if (allCells.Any(x => x.DataType == DataTypes.SharedString))
		tbpart = workbookPart.AddNewPart<SharedStringTablePart>();
	  for (uint i = 0; i < exlsSheetData.Count; i++)
	  {
		WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

		if (exlsSheetData[(int)i] is ExlsSheetData xlsxSheetData)
		{
		  worksheetPart.Worksheet = new Worksheet();
		  if (customFormat)
		  {
			var cols = GenerateColumns(xlsxSheetData.AllCells);
			if (cols != null)
			  worksheetPart.Worksheet.Append(cols);
		  }
		  else
		  {
			foreach (var c in xlsxSheetData.AllCells)
			  c.FormatIndex = c.RowIndex.Equals(1) ? 1u : 2u;
		  }

		  worksheetPart.Worksheet.Append(CreateSheetData(xlsxSheetData.SheetRows, tbpart));
		  if (doMerge)
			DoMerge(xlsxSheetData.AllCells, worksheetPart.Worksheet);
		}

		UInt32Value id = UInt32Value.FromUInt32(i + 1);
		Sheet sheet = new() { Id = spreadsheetDoc.WorkbookPart?.GetIdOfPart(worksheetPart), SheetId = id, Name = shNames[i][..Math.Min(shNames[i].Length, 30)] };

		sheets?.AppendChild(sheet);
	  }
	  workbookPart.Workbook.Save();
	}

	#endregion
  }
}

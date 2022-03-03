using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CommonLibraryStandard
{
  /// <summary>
  /// 这边都是读取Excel到内存对象的方法
  /// </summary>
  public partial class ExcelWrapper : IExcelWrapper
  {
	public ExlsSheetData ReadExlsSheetDataFromXlsxFile(ReadXlsxFileParams rxfPrms)
	{
	  if (Path.GetExtension(rxfPrms.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
	  {
		using var stream = File.Open(rxfPrms.FileName, FileMode.Open, FileAccess.Read, FileShare.Read);
		rxfPrms.TheStream = stream;
		return ReadExlsSheetDataFromStream(rxfPrms);
	  }
	  else
		throw new NotSupportedException("文件类型不是xlsx");
	}

	public ExlsSheetData ReadExlsSheetDataFromStream(ReadXlsxFileParams rxfPrms)
	{
	  ExlsSheetData result = new ExlsSheetData();
	  using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(rxfPrms.TheStream, false);
	  WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
	  var theSheet = workbookPart?.Workbook.Sheets.ElementAt(rxfPrms.SheetIndex) as Sheet;
	  var workSheetPart = workbookPart?.GetPartById(theSheet?.Id) as WorksheetPart;
	  uint rowIndex = 1;
	  var rows = workSheetPart?.Worksheet.Descendants<Row>();
	  if (rows?.Count() >= rxfPrms.StartRowIndex)
	  {
		result.SheetRows = new List<SheetRowItem>();
		string cellValue = string.Empty;
		int realIndex;
		var stringTable = workbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
		foreach (Row r in rows.Skip(rxfPrms.StartRowIndex - 1))
		{
		  var lstDataItems = new List<SheetCellItem>();
		  uint colIndex = 1;
		  foreach (Cell theCell in r.Elements<Cell>())
		  {
			cellValue = string.Empty;
			if (theCell.InnerText.Length > 0)
			  cellValue = GetCellValue(stringTable, theCell);

			realIndex = CellReferenceToIndex(theCell);
			//empty cell was skpipped
			if (colIndex <= realIndex)
			  colIndex = (uint)realIndex + 1;

			lstDataItems.Add(new SheetCellItem { ColIndex = colIndex++, Data = cellValue, DataType = DataTypes.String });
		  }

		  result.SheetRows.Add(new SheetRowItem(lstDataItems, rowIndex++));
		}
	  }

	  return result;
	}

	private string GetCellValue(SharedStringTablePart stringTable, Cell theCell)
	{
	  string cellValue = theCell.InnerText;
	  if (theCell.DataType != null)
		switch (theCell.DataType.Value)
		{
		  case CellValues.SharedString:
			if (stringTable != null)
			  cellValue = stringTable.SharedStringTable.ElementAt(int.Parse(cellValue)).InnerText;
			break;
		  case CellValues.Boolean:
			cellValue = cellValue switch
			{
			  "0" => "FALSE",
			  _ => "TRUE",
			};
			break;
		}
	  else if (theCell.CellFormula != null)
		cellValue = theCell.CellValue?.InnerText;

	  return cellValue;
	}

	private int CellReferenceToIndex(Cell cell)
	{
	  int index = 0;
	  string reference = cell.CellReference?.ToString()?.ToUpper();

	  if (!string.IsNullOrWhiteSpace(reference))
		foreach (char ch in reference)
		  if (char.IsLetter(ch))
		  {
			int value = ch - 'A';
			index = (index == 0) ? value : ((index + 1) * 26) + value;
		  }
		  else
			return index;

	  return index;
	}

	//	/// <summary>
	//	/// xls转xlsx 
	//	/// </summary>
	//	/// <param name="filename"></param>
	//	/// <returns></returns>
	//	private string ConvertToXlsx(string filename)
	//	{
	//	  var xlApp = new Microsoft.Office.Interop.Excel.Application();
	//	  Microsoft.Office.Interop.Excel.Workbook xlWorkBook = null;
	//	  try
	//	  {
	//		xlWorkBook = xlApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
	//		xlApp.DisplayAlerts = false;
	//		var tempPath = Path.GetTempPath();

	//		filename = Path.Combine(tempPath, Path.GetFileNameWithoutExtension(filename) + Guid.NewGuid().ToString() + ".xlsx");
	//		xlWorkBook.SaveAs(filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value,
	//	Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, true, Missing.Value, Missing.Value, Missing.Value);
	//	  }
	//	  catch
	//	  {
	//		throw;
	//	  }
	//	  finally
	//	  {
	//		xlWorkBook?.Close();
	//		xlApp.Quit();
	//#pragma warning disable CA1416 // 验证平台兼容性
	//		_ = Marshal.ReleaseComObject(xlWorkBook);
	//		_ = Marshal.ReleaseComObject(xlApp);
	//#pragma warning restore CA1416 // 验证平台兼容性
	//	  }

	//	  return filename;
	//	}
  }
}
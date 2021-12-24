namespace CommonLibraryCore
{
  public class ExcelWrapper : IExcelWrapper
  {
    public ExcelWrapper()
    {
    }

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
      using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
      {
        WorkbookPart workbookPart = spreadsheetDoc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        Sheets sheets = spreadsheetDoc.WorkbookPart?.Workbook.AppendChild(new Sheets());
        SharedStringTablePart tbpart = null;

        var stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
        List<SheetCellItem> allCells = new List<SheetCellItem>();
        for (uint i = 0; i < exlsSheetData.Count; i++)
        {
          ExlsSheetData xlsxSheetData = exlsSheetData[(int)i] as ExlsSheetData;

          if (xlsxSheetData != null)
          {
            allCells.AddRange(xlsxSheetData.AllCells);
          }
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
          ExlsSheetData xlsxSheetData = exlsSheetData[(int)i] as ExlsSheetData;

          if (xlsxSheetData != null)
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
          Sheet sheet = new Sheet { Id = spreadsheetDoc.WorkbookPart?.GetIdOfPart(worksheetPart), SheetId = id, Name = shNames[i].Substring(0, Math.Min(shNames[i].Length, 30)) };

          sheets?.AppendChild(sheet);
        }
        workbookPart.Workbook.Save();
      }
    }

    public MemoryStream GenerateXlsxFile(List<BaseSheetData> exlsSheetData, string[] shNames, bool doMerge = false, bool customFormat = false)
    {
      MemoryStream ms = new MemoryStream();
      using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
      {
        WorkbookPart workbookPart = spreadsheetDoc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        Sheets sheets = spreadsheetDoc.WorkbookPart?.Workbook.AppendChild(new Sheets());
        SharedStringTablePart tbpart = null;

        var stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
        List<SheetCellItem> allCells = new List<SheetCellItem>();
        for (uint i = 0; i < exlsSheetData.Count; i++)
        {
          ExlsSheetData xlsxSheetData = exlsSheetData[(int)i] as ExlsSheetData;

          if (xlsxSheetData != null)
          {
            allCells.AddRange(xlsxSheetData.AllCells);
          }
        }
        if (allCells.Any())
        {
          var ss = customFormat ? GenerateStylesheet(allCells) : DefaultStylesheet();
          if (ss == null)
            ss = DefaultStylesheet();
          stylePart.Stylesheet = ss;
          stylePart.Stylesheet.Save();
        }

        if (allCells.Any(x => x.DataType == DataTypes.SharedString))
          tbpart = workbookPart.AddNewPart<SharedStringTablePart>();
        for (uint i = 0; i < exlsSheetData.Count; i++)
        {
          WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
          ExlsSheetData xlsxSheetData = exlsSheetData[(int)i] as ExlsSheetData;

          if (xlsxSheetData != null)
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
          Sheet sheet = new Sheet { Id = spreadsheetDoc.WorkbookPart?.GetIdOfPart(worksheetPart), SheetId = id, Name = shNames[i].Substring(0, Math.Min(shNames[i].Length, 30)) };

          sheets?.AppendChild(sheet);
        }
        workbookPart.Workbook.Save();
      }
      return ms;
    }

    private Cell CreateCell(string col, uint row, string text, DataTypes dataType, CellTextPart[] textParts, SharedStringTablePart shareStringPart, uint formatIndex = 0)
    {
      switch (dataType)
      {
        case DataTypes.Number:
          {
            return CreateNumberCell(col, row, text, formatIndex);
          }
        case DataTypes.String:
          {
            return CreateTextCell(col, row, text, formatIndex);
          }
        case DataTypes.SharedString:
          {
            return CreateSharedStringCell(col, row, formatIndex, textParts, shareStringPart);
          }
        default:
          {
            return null;
          }
      }
    }

    private Cell CreateSharedStringCell(string col, uint row, uint formatIndex, CellTextPart[] textParts, SharedStringTablePart shareStringPart)
    {
      Cell cell = new Cell { DataType = CellValues.SharedString, CellReference = col + row.ToString() };
      if (formatIndex > 0)
        cell.StyleIndex = formatIndex;

      var index = GenerateSharedStringItem(textParts, shareStringPart);
      cell.CellValue = new CellValue(index.ToString());
      return cell;
    }

    private Cell CreateNumberCell(string col, uint row, string v, uint formatIndex)
    {
      var result = new Cell { DataType = CellValues.Number, CellReference = col + row.ToString(), CellValue = new CellValue(v) };
      if (formatIndex > 0)
        result.StyleIndex = formatIndex;

      return result;
    }

    private Cell CreateTextCell(string col, uint row, string text, uint formatIndex)
    {
      Cell cell = new Cell { DataType = CellValues.InlineString, CellReference = col + row.ToString() };
      if (formatIndex > 0)
        cell.StyleIndex = formatIndex;

      InlineString istring = new InlineString();
      Text t = new Text { Text = string.IsNullOrEmpty(text) ? string.Empty : text };
      istring.AppendChild(t);
      cell.AppendChild(istring);
      return cell;
    }

    private string ColumnLetter(int intCol)
    {
      int intFirstLetter = ((intCol - 1) / 676) + 64;
      int intSecondLetter = (((intCol - 1) % 676) / 26) + 64;
      int intThirdLetter = ((intCol - 1) % 26) + 65;

      char firstLetter = (intFirstLetter > 64) ? (char)intFirstLetter : ' ';
      char secondLetter = (intSecondLetter > 64) ? (char)intSecondLetter : ' ';
      char thirdLetter = (char)intThirdLetter;

      return string.Concat(firstLetter, secondLetter, thirdLetter).Trim();
    }

    /// <summary>
    /// 生成一个sheet数据
    /// </summary>
    /// <param name="rowItems"></param>
    /// <returns></returns>
    private SheetData CreateSheetData(IEnumerable<SheetRowItem> rowItems, SharedStringTablePart sharedStringPart)
    {
      SheetData sheetData = null;

      if (rowItems != null)
      {
        sheetData = new SheetData();

        if (rowItems.Any())
        {
          foreach (var r in rowItems.OrderBy(x => x.RowIndex))
          {
            sheetData.AppendChild(CreateSheetRow(r, r.RowIndex, sharedStringPart));
          }
        }
      }

      return sheetData;
    }
    /// <summary>
    /// 合并单元格
    /// </summary>
    /// <param name="cellItems"></param>
    /// <param name="worksheet"></param>
    private void DoMerge(IEnumerable<SheetCellItem> cellItems, Worksheet worksheet)
    {
      MergeCells mergeCells = new MergeCells();
      if (cellItems.Any(x => x.MergeToColIndex > 0 || x.MergeToRowIndex > 0))
      {
        foreach (var itm in cellItems.Where(x => x.MergeToColIndex > 0 || x.MergeToRowIndex > 0))
        {
          var mergeFrom = ColumnLetter((int)itm.ColIndex) + itm.RowIndex.ToString();
          var mergeTo = ColumnLetter((int)(itm.MergeToColIndex < itm.ColIndex ? itm.ColIndex : itm.MergeToColIndex)) + (itm.MergeToRowIndex < itm.RowIndex ? itm.RowIndex : itm.MergeToRowIndex).ToString();
          MergeCell mergeCell = new MergeCell { Reference = new StringValue(mergeFrom + ":" + mergeTo) };
          mergeCells.Append(mergeCell);
        }
      }

      if (mergeCells.Any())
      {
        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
      }
    }
    private int GenerateSharedStringItem(CellTextPart[] textParts, SharedStringTablePart shareStringPart)
    {
      if (shareStringPart.SharedStringTable == null)
        shareStringPart.SharedStringTable = new SharedStringTable();

      SharedStringItem ssItm = new SharedStringItem();

      foreach (var part in textParts)
      {
        Run run = new Run();
        var txt = part.Text;
        if (part.PartFormat != null)
        {
          RunProperties runProperties = new RunProperties();
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
        Text text = new Text { Text = txt };
        run.Append(text);

        ssItm.Append(run);
      }
      // The text does not exist in the part. Create the SharedStringItem and return its index.
      shareStringPart.SharedStringTable.AppendChild(ssItm);
      shareStringPart.SharedStringTable.Save();

      // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
      //foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
      //{
      //  if (item.InnerText.Equals(ssItm.InnerText))
      //    return i;

      //  i++;
      //}
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
      if (cellItems != null && cellItems.Any(x => x.CellFormats != null))
      {
        IEnumerable<SheetCellFormats> cfs = cellItems.Where(x => x.CellFormats != null).Select(x => x.CellFormats).Distinct();
        Fills fills = new Fills(new Fill(new PatternFill() { PatternType = PatternValues.None }),
                            new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }));
        Dictionary<string, Fill> fgColorAndFills = new Dictionary<string, Fill>();
        if (cfs.Any(x => !string.IsNullOrEmpty(x.FGColor)))
        {
          foreach (var fg in cfs.Where(x => !string.IsNullOrEmpty(x.FGColor)).Select(x => x.FGColor).Distinct())
          {
            var f = new Fill(new PatternFill() { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor() { Rgb = new HexBinaryValue(fg) } });
            fgColorAndFills.Add(fg, f);
            fills.AppendChild(f);
          }
        }
        Fonts fonts = new Fonts(new Font(new FontSize() { Val = 10 }),
                                                        new Font(new FontSize() { Val = 12 }, new Bold() { Val = BooleanValue.FromBoolean(true) }),
                                                        new Font(new FontSize() { Val = 10 }, new FontName() { Val = new StringValue("Arial") }));
        Borders borders = new Borders(new Border(),
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
          CellFormat cf = new CellFormat();
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
              CellFormat fc = new CellFormat();
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
        CellFormats cellFormats = new CellFormats(lstCellFormats);
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
      Fonts fonts = new Fonts(new Font(new FontSize() { Val = 10 }, new FontName { Val = "宋体" }),
          new Font(new FontSize() { Val = 10 }, new Bold() { Val = true }, new Color { Rgb = "993366" }, new FontName { Val = "宋体" }));
      Fills fills = new Fills(new Fill(new PatternFill() { PatternType = PatternValues.None }),
          new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }),
          new Fill(new PatternFill() { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor() { Rgb = new HexBinaryValue("C0C0C0") } }));
      Borders borders = new Borders(new Border(),
          new Border(new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                          new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                          new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                          new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                          new DiagonalBorder()));
      CellFormat headFormat = new CellFormat { FontId = 1, ApplyFont = true, FillId = 2, BorderId = 1, ApplyBorder = true };
      CellFormat dataFormat = new CellFormat { FontId = 0, ApplyFont = true, FillId = 0, BorderId = 1, ApplyBorder = true };

      CellFormats cellFormats = new CellFormats(new CellFormat() { FillId = 0, FontId = 0 }, headFormat, dataFormat);
      return new Stylesheet(fonts, fills, borders, cellFormats);
    }

    #endregion

    private Columns GenerateColumns(IEnumerable<SheetCellItem> sheetCells)
    {
      Columns result = null;
      if (sheetCells != null && sheetCells.Any(x => x.CustWidth > 0))
      {
        result = new Columns();
        foreach (var itm in sheetCells.Where(x => x.CustWidth > 0).OrderBy(x => x.ColIndex))
          result.Append(new Column { CustomWidth = true, Min = itm.ColIndex, Max = itm.ColIndex, Width = itm.CustWidth });
      }
      return result;
    }
    protected Row CreateSheetRow(SheetRowItem item, uint rowIndex, SharedStringTablePart shareStringPart)
    {
      Row result = null;
      if (item != null && item.RowCells != null && item.RowCells.Any())
      {
        result = new Row { RowIndex = rowIndex };
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
    #endregion

    #region 读取Excel相关

    public ExlsSheetData ReadSpreadSheetDoc(string filename, int sheetIndex, int startRowIndex)
    {
      if (Path.GetExtension(filename).Equals(".xls", StringComparison.OrdinalIgnoreCase))
        filename = ConvertToXlsx(filename);

      using (var stream = File.Open(filename, FileMode.Open))
      {
        return ReadSpreadSheetDoc(stream, sheetIndex, startRowIndex);
      }
    }

    public ExlsSheetData ReadSpreadSheetDoc(Stream stream, int sheetIndex, int startRowIndex)
    {
      ExlsSheetData result = new ExlsSheetData();
      using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
      {
        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
        var theSheet = workbookPart?.Workbook.Sheets.ElementAt(sheetIndex) as Sheet;
        var workSheetPart = workbookPart?.GetPartById(theSheet?.Id) as WorksheetPart;
        string cellValue = string.Empty;
        uint rowIndex = 1;
        var lstDataRows = new List<SheetRowItem>();
        var rows = workSheetPart?.Worksheet.Descendants<Row>();
        if (rows.Count() >= startRowIndex)
        {
          foreach (Row r in rows.Skip(startRowIndex - 1))
          {
            var lstDataItems = new List<SheetCellItem>();
            uint colIndex = 0;
            foreach (Cell theCell in r.Elements<Cell>())
            {
              if (theCell.InnerText.Length > 0)
                cellValue = GetCellValue(workbookPart, theCell);
              else
                cellValue = string.Empty;
              var realIndex = CellReferenceToIndex(theCell);
              if (colIndex < realIndex)
              {
                while (colIndex < realIndex)
                  lstDataItems.Add(new SheetCellItem { ColIndex = ++colIndex, Data = string.Empty, DataType = DataTypes.String });
              }//empty cell was skpipped

              lstDataItems.Add(new SheetCellItem { ColIndex = ++colIndex, Data = cellValue, DataType = DataTypes.String });
            }

            lstDataRows.Add(new SheetRowItem(lstDataItems, rowIndex));
            rowIndex++;
          }
        }
        result.SheetRows = lstDataRows;
      }

      return result;
    }

    private string GetCellValue(WorkbookPart workbookPart, Cell theCell)
    {
      string cellValue = theCell.InnerText;
      if (theCell.DataType != null)
      {
        switch (theCell.DataType.Value)
        {
          case CellValues.SharedString:
            var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (stringTable != null)
            {
              cellValue = stringTable.SharedStringTable.ElementAt(int.Parse(cellValue)).InnerText;
            }
            break;
          case CellValues.Boolean:
            switch (cellValue)
            {
              case "0":
                cellValue = "FALSE";
                break;
              default:
                cellValue = "TRUE";
                break;
            }
            break;
        }
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
        {
          if (char.IsLetter(ch))
          {
            int value = ch - 'A';
            index = (index == 0) ? value : ((index + 1) * 26) + value;
          }
          else
          {
            return index;
          }
        }
      return index;
    }

    private string ConvertToXlsx(string filename)
    {
      var xlApp = new Microsoft.Office.Interop.Excel.Application();
      Microsoft.Office.Interop.Excel.Workbook xlWorkBook = null;
      try
      {
        xlWorkBook = xlApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        xlApp.DisplayAlerts = false;
        var tempPath = Path.GetTempPath();

        filename = Path.Combine(tempPath, Path.GetFileNameWithoutExtension(filename) + Guid.NewGuid().ToString() + ".xlsx");
        xlWorkBook.SaveAs(filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value,
    Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, true, Missing.Value, Missing.Value, Missing.Value);
      }
      catch
      {
        throw;
      }
      finally
      {
        xlWorkBook?.Close();
        xlApp.Quit();
#pragma warning disable CA1416 // 验证平台兼容性
        _ = Marshal.ReleaseComObject(xlWorkBook);
        _ = Marshal.ReleaseComObject(xlApp);
#pragma warning restore CA1416 // 验证平台兼容性
      }

      return filename;
    }
    #endregion
  }
}

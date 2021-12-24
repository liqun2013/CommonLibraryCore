namespace CommonLibraryCore
{
  public abstract class BaseSheetData
  {
  }

  public class ExlsSheetData : BaseSheetData
  {
    public ExlsSheetData()
    { }
    public ExlsSheetData(List<SheetRowItem> rowItems)
    {
      SheetRows = rowItems;
    }
    public List<SheetCellItem> AllCells
    {
      get
      {
        List<SheetCellItem> cells = new List<SheetCellItem>();
        if (SheetRows != null && SheetRows.Any())
          foreach (var itm in SheetRows)
            cells.AddRange(itm.RowCells);
        return cells;
      }
    }
    public List<SheetRowItem> SheetRows { get; set; }
    public void AddRow(SheetRowItem row)
    {
      if (SheetRows == null)
        SheetRows = new List<SheetRowItem>();

      if (SheetRows.Any(x => x.RowIndex == row.RowIndex))
        throw new Exception("rowindex exist");

      SheetRows.Add(row);
    }
    public void AddCell(SheetCellItem cell, uint rindex)
    {
      AddCell(cell, rindex, null);
    }
    public void AddCell(SheetCellItem cell, uint rindex, SheetRowFormats rowFormats)
    {
      if (rindex < 1)
        throw new ArgumentOutOfRangeException("rindex", "rowindex must greater than zero");

      var r = FindRow(rindex);
      if (r == null)
      {
        r = new SheetRowItem(new List<SheetCellItem>(), rindex);
        if (rowFormats != null)
          r.RowHeight = rowFormats.RowHeight;
        AddRow(r);
      }
      cell.RowIndex = rindex;
      r.RowCells.Add(cell);
    }

    protected SheetRowItem FindRow(uint rindex)
    {
      SheetRowItem result = null;

      if (SheetRows != null && SheetRows.Any(x => x.RowIndex == rindex))
        result = SheetRows.First(x => x.RowIndex == rindex);

      return result;
    }
    public List<T> ToList<T>() where T : new()
    {
      List<T> result = new List<T>();

      if (SheetRows != null && SheetRows.Any())
      {
        foreach (SheetRowItem itm in SheetRows)
        {
          T obj = new T();
          var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
          for (var i = 0; i < itm.RowCells.Count; i++)
          {
            var property = FindProperty(properties, i + 1);
            if (property != null)
              SetPropertyValue(property, obj, itm.RowCells[i].Data);
          }
          result.Add(obj);
        }
      }

      return result;
    }
    private PropertyInfo FindProperty(PropertyInfo[] properties, int order)
    {
      foreach (var itm in properties)
      {
        var attrs = itm.GetCustomAttributes(false);
        if (attrs != null && attrs.Any() && attrs.Any(x => x.GetType() == typeof(ColAttribute)))
        {
          var rowDataAttr = attrs.FirstOrDefault(x => x.GetType() == typeof(ColAttribute)) as ColAttribute;
          if (rowDataAttr?.IsImport == true && rowDataAttr.OrderInImporter == order)
            return itm;
        }
        //else
        //    throw new Exception("ColAttribute not found");
      }

      return null;
    }

    private void SetPropertyValue<T>(PropertyInfo p, T obj, string v)
    {
      switch (p.PropertyType.Name.ToLower())
      {
        case "string":
          {
            p.SetValue(obj, v, null);
          }
          break;
        case "datetime":
          {
            p.SetValue(obj, v.ToDateTime(false), null);
          }
          break;
        case "int32":
          {
            p.SetValue(obj, v.ToInt32(false), null);
          }
          break;
        case "int64":
          {

            p.SetValue(obj, v.ToInt64(false), null);
          }
          break;
        case "boolean":
          {
            p.SetValue(obj, v.ToBoolean(false), null);
          }
          break;
        case "double":
          {
            p.SetValue(obj, v.ToDouble(false), null);
          }
          break;
        case "decimal":
          {
            p.SetValue(obj, v.ToDecimal(false), null);
          }
          break;
        case "float":
          {
            p.SetValue(obj, v.ToFloat(false), null);
          }
          break;
      }
    }

  }

  public class SheetRowItem
  {
    public SheetRowItem(List<SheetCellItem> rowCells, uint rindex)
    {
      RowCells = rowCells;
      RowIndex = rindex;
    }
    public uint RowIndex { get; set; }
    public uint RowHeight { get; set; }
    public List<SheetCellItem> RowCells { get; set; }
  }

  /// <summary>
  /// 定义Excel单元格相关属性
  /// </summary>
  public class SheetCellItem
  {
    public uint RowIndex { get; set; }
    public uint ColIndex { get; set; }
    public uint MergeToRowIndex { get; set; }
    public uint MergeToColIndex { get; set; }
    public string Data { get; set; }
    public DataTypes DataType { get; set; }
    public SheetCellFormats CellFormats { get; set; }
    public uint FormatIndex { get; set; }
    public uint CustWidth { get; set; }
    public uint CustHeight { get; set; }
    public CellTextPart[] Texts { get; set; }
  }

  public class SheetRowFormats : IEquatable<SheetRowFormats>
  {
    public uint RowHeight { get; set; }
    public bool Equals(SheetRowFormats other)
    {
      if (ReferenceEquals(other, null))
        return false;
      if (ReferenceEquals(this, other))
        return true;

      return RowHeight.Equals(other.RowHeight);
    }
    public override int GetHashCode()
    {
      return RowHeight.GetHashCode();
    }

    public override bool Equals(object obj)
    {
      return Equals(obj as SheetRowFormats);
    }
  }
  public class SheetCellFormats : IEquatable<SheetCellFormats>
  {
    public SheetCellFormats()
    {
      FontSize = 0;
      FontBold = false;
      FontColor = string.Empty;
      FGColor = string.Empty;
      FontName = string.Empty;
      CellWidth = 0;
      CellHeight = 0;
      HorizontalAlignment = HorizontalAlignments.Default;
      VerticalAlignment = VerticalAlignments.Default;
      Borders = new bool[4];
    }
    public double FontSize { get; set; }
    public bool FontBold { get; set; }
    public string FontName { get; set; }
    public string FontColor { get; set; }
    public string FGColor { get; set; }
    public bool[] Borders { get; set; }
    public int CellWidth { get; set; }
    public int CellHeight { get; set; }
    public bool WrapText { get; set; }
    public HorizontalAlignments HorizontalAlignment { get; set; }
    public VerticalAlignments VerticalAlignment { get; set; }
    public bool Equals(SheetCellFormats other)
    {
      if (ReferenceEquals(other, null))
        return false;

      if (ReferenceEquals(this, other))
        return true;

      return FontSize.Equals(other.FontSize) && FontName.Equals(other.FontName) && FontBold.Equals(other.FontBold) && FontColor.Equals(other.FontColor) && FGColor.Equals(other.FGColor) && Borders[0].Equals(other.Borders[0]) && Borders[1].Equals(other.Borders[1]) && Borders[2].Equals(other.Borders[2]) && Borders[3].Equals(other.Borders[3]) && CellWidth.Equals(other.CellWidth) && CellHeight.Equals(other.CellHeight) && HorizontalAlignment == other.HorizontalAlignment && VerticalAlignment == other.VerticalAlignment && WrapText == other.WrapText;
    }
    public override int GetHashCode()
    {
      return FontSize.GetHashCode() ^ FontName.GetHashCode() ^ FontBold.GetHashCode() ^ FontColor.GetHashCode() ^ FGColor.GetHashCode() ^ Borders[0].GetHashCode() ^ Borders[1].GetHashCode() ^ Borders[2].GetHashCode() ^ Borders[3].GetHashCode() ^ CellWidth.GetHashCode() ^ CellHeight.GetHashCode() ^ HorizontalAlignment.GetHashCode() ^ VerticalAlignment.GetHashCode() ^ WrapText.GetHashCode();
    }
  }
  public class CellTextPart
  {
    public string Text { get; set; }
    public DataTypes TheDataType { get; set; }
    public SheetCellFormats PartFormat { get; set; }
  }
  public enum HorizontalAlignments
  {
    Default,
    Left,
    Center,
    Right
  }
  public enum VerticalAlignments
  {
    Default,
    Top,
    Middle,
    Bottom
  }
}

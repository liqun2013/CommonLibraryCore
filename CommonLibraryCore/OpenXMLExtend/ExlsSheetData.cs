namespace CommonLibraryCore;
public class ExlsSheetData
{
  public ExlsSheetData() { }
  public ExlsSheetData(List<SheetRowItem> rowItems, StyleItem headStyle, StyleItem dataStyle)
  {
	SheetRows = rowItems;
	HeadStyle = headStyle;
	DataStyle = dataStyle;
  }
  public List<SheetRowItem> SheetRows { get; set; }
  /// <summary>
  /// 表头统一样式
  /// </summary>
  public StyleItem HeadStyle { get; set; }
  /// <summary>
  /// 数据的统一样式
  /// </summary>
  public StyleItem DataStyle { get; set; }
  public void AddRow(SheetRowItem row)
  {
	if (SheetRows == null)
	  SheetRows = new List<SheetRowItem>();

	if (SheetRows.Any(x => x.RowIndex == row.RowIndex))
	  throw new Exception("rowindex exist");

	SheetRows.Add(row);
  }
  public SheetRowItem FirstRow => SheetRows?.OrderBy(x => x.RowIndex)
										 .FirstOrDefault();

  public List<SheetCellItem> AllCells => SheetRows?.SelectMany(x => x.RowCells).ToList();
  /// <summary>
  /// 转成List数据
  /// </summary>
  /// <typeparam name="T"></typeparam>
  public List<T> ToList<T>() where T : new()
  {
	List<T> result = new();

	if (SheetRows?.Any() == true)
	{
	  var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty);
	  List<(uint order, PropertyInfo property)> lstOrderAndProperty = new();
	  foreach (var itm in properties)
	  {
		var attr = itm.GetCustomAttribute<ColAttribute>(false);
		if (attr != null && !attr.NotImport)
		  lstOrderAndProperty.Add((attr.OrderInImporter, itm));
	  }
	  foreach (SheetRowItem itm in SheetRows)
	  {
		T obj = new();
		for (var i = 0; i < itm.RowCells.Count; i++)
		{
		  if (lstOrderAndProperty.Any(x => x.order.Equals(itm.RowCells[i].ColIndex) && x.property != null))
		  {
			PropertyInfo property = lstOrderAndProperty.FirstOrDefault(x => x.order.Equals(itm.RowCells[i].ColIndex)).property; //FindProperty(properties, itm.RowCells[i].ColIndex);
			SetPropertyValue(property, obj, itm.RowCells[i].Data);
		  }
		}
		result.Add(obj);
	  }
	}

	return result;
  }

  public ExlsSheetData ToSheetdata<T>(IEnumerable<T> data, bool addHeadRow)
  {
	if (data?.Any() == true)
	{
	  uint rowIndex = 1;
	  uint colIndex = 1;
	  var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty);
	  //下面提取数据时根据这里的DisplayOrder顺序转换成column,根据PropertyName提取T中的属性值
	  List<(uint DisplayOrder, string ColName, string PropertyName, uint ColWidth)> rowDataAtts = new();
	  SheetRows = new List<SheetRowItem>();

	  if (addHeadRow)
	  {
		var headRow = new SheetRowItem(new List<SheetCellItem>(), rowIndex++);

		foreach (MemberInfo itm in properties)
		{//存在有的属性定义了ColAttribute,有的没定义的情况
		  ColAttribute colAtt = itm.GetCustomAttribute<ColAttribute>();
		  if (colAtt != null)
		  {
			if (!colAtt.NotExport)
			{
			  rowDataAtts.Add((colAtt.DisplayOrder, string.IsNullOrWhiteSpace(colAtt.ColName) ? itm.Name : colAtt.ColName, itm.Name, colAtt.ColWidth));
			  colIndex = colAtt.DisplayOrder;
			}
		  }
		  else
		  {
			if (rowDataAtts.Any())
			  colIndex = rowDataAtts.OrderBy(x => x.DisplayOrder).Last().DisplayOrder + 1;
			rowDataAtts.Add((colIndex, itm.Name, itm.Name, uint.MinValue));
		  }
		}

		colIndex = 1;
		foreach ((_, string ColName, _, uint ColWidth) in rowDataAtts.OrderBy(x => x.DisplayOrder))
		  headRow.RowCells.Add(new() { Data = ColName, ColIndex = colIndex++, DataType = DataTypes.String, CustWidth = ColWidth });

		SheetRows.Add(headRow); //the head row added
	  }//添加第一行为标题行
	  else
	  {
		foreach (MemberInfo itm in properties)
		{
		  if (rowDataAtts.Any())
			colIndex = rowDataAtts.OrderBy(x => x.DisplayOrder).Last().DisplayOrder + 1;
		  rowDataAtts.Add((colIndex, itm.Name, itm.Name, uint.MinValue));
		}
	  }

	  foreach (T itm in data)
	  {
		colIndex = 1;
		var dataRow = new SheetRowItem(new List<SheetCellItem>(), rowIndex);
		foreach ((_, _, string PropertyName, _) in rowDataAtts.OrderBy(x => x.DisplayOrder))
		{
		  var theProperty = properties.FirstOrDefault(x => x.Name.Equals(PropertyName));
		  if (theProperty != null)
		  {
			object v = theProperty.GetValue(itm);
			var cellItm = new SheetCellItem
			{ ColIndex = colIndex, RowIndex = rowIndex, Data = v?.ToString(true, string.Empty) ?? string.Empty, DataType = v?.IsNumeric() == true ? DataTypes.Number : DataTypes.String };
			dataRow.RowCells.Add(cellItm);
		  }
		  colIndex++;
		}
		rowIndex++;

		SheetRows.Add(dataRow);	//add a data row
	  }
	}
	return this;
  }

  public bool IsEmpty => AllCells?.Any() != true;
  public bool IsSharedStringExist => AllCells?.Any(x => x.DataType == DataTypes.SharedString) == true;
  private void SetPropertyValue<T>(PropertyInfo p, T obj, string v)
  {
	switch (p.PropertyType.Name.ToLower())
	{
	  case "string":
		p.SetValue(obj, v, null);
		break;
	  case "datetime":
		p.SetValue(obj, v.ToDateTime(false), null);
		break;
	  case "int32":
		p.SetValue(obj, v.ToInt32(false), null);
		break;
	  case "int64":
		p.SetValue(obj, v.ToInt64(false), null);
		break;
	  case "boolean":
		p.SetValue(obj, v.ToBoolean(false), null);
		break;
	  case "double":
		p.SetValue(obj, v.ToDouble(false), null);
		break;
	  case "decimal":
		p.SetValue(obj, v.ToDecimal(false), null);
		break;
	  case "float":
		p.SetValue(obj, v.ToFloat(false), null);
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
  /// <summary>
  /// 行索引，以1开始
  /// </summary>
  public uint RowIndex { get; set; }
  /// <summary>
  /// 行高只需要设置在行上，不需要设置每个单元格
  /// </summary>
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
  public StyleItem CellStyle { get; set; }
  public uint StyleIndex { get; set; }
  public uint CustWidth { get; set; }
  public uint CustHeight { get; set; }
  public CellTextPart[] Texts { get; set; }
}

public class CellTextPart
{
  public string Text { get; set; }
  public DataTypes TheDataType { get; set; }
  public FontStyle PartFontStyle { get; set; }
}

namespace CommonLibraryCore;
/// <summary>
/// 调用写入Excel文件的参数
/// </summary>
public sealed class WriteXlsxFileParams
{
  private IEnumerable<string> sheetNames;

  public WriteXlsxFileParams(IEnumerable<ExlsSheetData> lstSheetData, IEnumerable<string> lstSheetName, string fn, bool doMerge, bool custFormat)
  {
	SheetData = lstSheetData;
	sheetNames = lstSheetName;
	FileName = fn;
	DoMerge = doMerge;
	CustFormat = custFormat;
  }
  public WriteXlsxFileParams(IEnumerable<ExlsSheetData> lstSheetData, IEnumerable<string> lstSheetName, bool doMerge = false, bool custFormat = false)
	: this(lstSheetData, lstSheetName, string.Empty, doMerge, custFormat)
  { }
  public WriteXlsxFileParams(ExlsSheetData sheetData, string sheetName, string fn, bool doMerge = false, bool custFormat = false)
	: this(new List<ExlsSheetData> { sheetData }, new List<string> { sheetName }, fn, doMerge, custFormat)
  { }
  public WriteXlsxFileParams(ExlsSheetData sheetData, string sheetName, bool doMerge = false, bool custFormat = false)
	: this(sheetData, sheetName, string.Empty, doMerge, custFormat)
  { }
  public WriteXlsxFileParams(WriteXlsxFileParams wrtXFPrms)
	: this(wrtXFPrms.SheetData, wrtXFPrms.SheetNames, wrtXFPrms.FileName, wrtXFPrms.DoMerge, wrtXFPrms.CustFormat)
  { }
  /// <summary>
  /// sheet data集合
  /// </summary>
  public IEnumerable<ExlsSheetData> SheetData { get; set; }
  /// <summary>
  /// sheet名称集合,名称长度限制在30位,超过30的将被截断
  /// </summary>
  public IEnumerable<string> SheetNames => sheetNames?.Select(x => x[..Math.Min(x?.Length ?? 0, 30)]);
  /// <summary>
  /// 写入的文件名
  /// </summary>
  public string FileName { get; set; }
  /// <summary>
  /// 是否需要合并单元格
  /// </summary>
  public bool DoMerge { get; set; }
  /// <summary>
  /// 是否需要自定义格式
  /// </summary>
  public bool CustFormat { get; set; }
}

/// <summary>
/// 调用读取Excel文件的参数
/// </summary>
public sealed class ReadXlsxFileParams
{
  private ReadXlsxFileParams(FileStream sm, string fn, int shIndex, int startRowIndex)
  {
	TheStream = sm;
	FileName = fn;
	SheetIndex = shIndex;
	StartRowIndex = startRowIndex;
  }
  public ReadXlsxFileParams(string fn, int shIndex = 0, int startRowIndex = 1)
	: this(null, fn, shIndex, startRowIndex)
  { }
  public ReadXlsxFileParams(FileStream sm, int shIndex = 0, int startRowIndex = 1)
	: this(sm, string.Empty, shIndex, startRowIndex)
  { }
  /// <summary>
  /// 从流中读取
  /// </summary>
  public FileStream TheStream { get; set; }
  /// <summary>
  /// 从指定文件读取
  /// </summary>
  public string FileName { get; set; }
  /// <summary>
  /// 要读取Excel文件中的第几个sheet,从0开始
  /// </summary>
  public int SheetIndex { get; set; }
  /// <summary>
  /// 要从第几行开始读取,从1开始
  /// </summary>
  public int StartRowIndex { get; set; }
}

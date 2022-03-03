using System;

namespace CommonLibraryStandard
{
  [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
  public class ColAttribute : Attribute
  {
	/// <summary>
	/// 设置是否导出该字段
	/// </summary>
	public bool NotExport { get; set; }
	/// <summary>
	/// 列的顺序
	/// </summary>
	public uint DisplayOrder { get; set; }
	/// <summary>
	/// 列宽
	/// </summary>
	public uint ColWidth { get; set; }
	public DataTypes DataTxtType { get; set; }
	public string ColName { get; set; }
	public ColAttribute()
	{
	  DataTxtType = DataTypes.String;
	  ColName = string.Empty;
	}

	#region 导入相关的
	/// <summary>
	/// 是否导入该字段
	/// </summary>
	public bool NotImport { get; set; }
	/// <summary>
	/// 设置该字段在导入的文件（Excel）第几列
	/// </summary>
	public uint OrderInImporter { get; set; }
	#endregion
  }
}
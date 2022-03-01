namespace CommonLibraryCoreTests;

internal class ExportTestData
{
	[Col(ColName = "TenantId", DisplayOrder = 1)]
	public int TenantId { get; set; }
	[Col(ColName = "", DisplayOrder = 2)]
	public long ShelfId { get; set; }
	[Col(ColName = "", DisplayOrder = 3)]
	public long LocationCategoryId { get; set; }

	[Col(ColName = "", DisplayOrder = 4)]
	public string Code { get; set; }

	[Col(ColName = "", DisplayOrder = 5)]
	public string Name { get; set; }

	[Col(ColName = "", DisplayOrder = 6, ColWidth = 20)]
	public string Description { get; set; }

	[Col(ColName = "对齐", DisplayOrder = 7)]
	public int Orientation { get; set; }
	public int TheLevel { get; set; }
	[Col(ColName = "", DisplayOrder = 8)]

	public int TheColumn { get; set; }
	[Col(ColName = "", DisplayOrder = 9)]

	public int TheRow { get; set; }
	[Col(ColName = "", DisplayOrder = 10)]

	public decimal Length { get; set; }
	[Col(ColName = "", DisplayOrder = 11)]

	public decimal Width { get; set; }
	[Col(ColName = "", DisplayOrder = 12)]

	public decimal Height { get; set; }
	[Col(ColName = "", DisplayOrder = 13)]

	public decimal Volume { get; set; }
	[Col(ColName = "", DisplayOrder = 14)]

	public int Status { get; set; }

	public bool IsDeleted { get; set; }
	public DateTime CreateTime { get; set; }

	public string Creator { get; set; }
	public DateTime? LastUpdateTime { get; set; }
	public decimal? HourlyUnitFee { get; set; }
	public int? MaxStorageTime { get; set; }
	public int? UnitFee { get; set; }

	public string BoardId { get; set; }
	public string BoxId { get; set; }
	public List<string> Packages { get; set; }
}

namespace CommonLibraryCoreTests;

internal class ImportTestData
{
	[Col(OrderInImporter = 1)]
	public int Id { get; set; }
	[Col(OrderInImporter = 2)]
	public DateTime Dt { get; set; }
	[Col(OrderInImporter = 3)]
	public string Dl { get; set; }
	[Col(OrderInImporter = 4)]
	public string Num { get; set; }
	[Col(OrderInImporter = 5)]
	public string Name { get; set; }
	[Col(OrderInImporter = 6)]
	public string Code { get; set; }
	[Col(OrderInImporter = 7)]
	public string Desc { get; set; }
	[Col(OrderInImporter = 8)]
	public string Mo { get; set; }
	[Col(OrderInImporter = 9)]
	public string Mobile { get; set; }
	[Col(OrderInImporter = 10)]
	public string Cus { get; set; }
	[Col(OrderInImporter = 11)]
	public string CusMobile { get; set; }
	[Col(OrderInImporter = 12)]
	public string Ac { get; set; }
	[Col(OrderInImporter = 13)]
	public string Tp { get; set; }
	[Col(OrderInImporter = 14)]
	public string Remark { get; set; }
	[Col(NotImport = true)]
	public string Remark2 { get; set; }
}
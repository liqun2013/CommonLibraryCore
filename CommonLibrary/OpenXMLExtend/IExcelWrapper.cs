using System.IO;
using System.Threading.Tasks;

namespace CommonLibraryStandard
{
  public interface IExcelWrapper
  {
	Task WriteToXlsxFileAsync(WriteXlsxFileParams wxfPrms);
	void WriteToXlsxFile(WriteXlsxFileParams wxfPrms);
	MemoryStream WriteToMemoryStream(WriteXlsxFileParams wxfPrms);
	ExlsSheetData ReadExlsSheetDataFromXlsxFile(ReadXlsxFileParams rxfPrms);
	ExlsSheetData ReadExlsSheetDataFromStream(ReadXlsxFileParams rxfPrms);
  }
}
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace ConvertToExcelLibrary
{
    public interface IAppProcessor
    {
        List<IMeasure> LoadMeasureFiles(List<string> paths);
        Workbook LoadTemplateFile(string path);
        LogMessage SaveMeasureFilesInExcel(IMeasureRepository repo, Workbook template);
    }
}
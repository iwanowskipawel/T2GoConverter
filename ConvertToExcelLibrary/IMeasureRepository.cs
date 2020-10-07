using System.Collections.Generic;

namespace ConvertToExcelLibrary
{
    public interface IMeasureRepository
    {
        List<IMeasure> Measures { get; set; }

        string GetMeasuresDirectory();
    }
}
using System.Collections.Generic;

namespace ConvertToExcelLibrary
{
    public interface IMeasure
    {
        short DragHi { get; set; }
        short DragLo { get; set; }
        List<short> Friction { get; set; }
        string OriginalFilePath { get; set; }

        string GetDirectory();
        List<double> GetFrictionFactors();
        string GetName();
    }
}
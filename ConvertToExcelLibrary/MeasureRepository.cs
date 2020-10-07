using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertToExcelLibrary
{
    public class MeasureRepository : IMeasureRepository
    {
        public List<IMeasure> Measures { get; set; }

        public string GetMeasuresDirectory()
        {
            return Measures.FirstOrDefault().GetDirectory();
        }
    }
}

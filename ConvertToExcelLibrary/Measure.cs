using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertToExcelLibrary
{
    public class Measure : IMeasure
    {
        public string OriginalFilePath { get; set; }
        public short DragHi { get; set; }
        public short DragLo { get; set; }
        public List<short> Friction { get; set; } = new List<short>();

        public string GetDirectory()
        {
            return $"{ Path.GetDirectoryName(OriginalFilePath) }";
        }

        public string GetName()
        {
            return Path.GetFileNameWithoutExtension(OriginalFilePath);
        }

        public List<double> GetFrictionFactors()
        {
            List<double> factors = new List<double>();

            foreach (var f in Friction)
            {
                factors.Add(CalculateFrictionFactor(f));
            }

            return factors;
        }

        double CalculateFrictionFactor(short friction)
        {
            double f = (double)(friction - DragLo) / (DragHi - DragLo);

            return Math.Round(f, 3);
        }
    }
}

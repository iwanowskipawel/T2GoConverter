using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ConvertToExcelLibrary
{
    public static class MeasureImporter
    {
        public static IMeasure Load(string path)
        {
            IMeasure measure = new Measure();

            XDocument xmlFile = XDocument.Load(path);
            measure.DragHi = short.Parse(xmlFile.Descendants("DragHi").First().Value);
            measure.DragLo = short.Parse(xmlFile.Descendants("DragLo").First().Value);

            foreach (var f in xmlFile.Descendants("Friction"))
            {
                measure.Friction.Add(short.Parse(f.Value));
            }

            measure.OriginalFilePath = path;

            return measure;
        }
    }
}

using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace ConvertToExcelLibrary
{
    public class AppProcessor : IAppProcessor
    {
        public List<IMeasure> LoadMeasureFiles(List<string> paths)
        {
            List<IMeasure> measures = new List<IMeasure>();

            foreach (var path in paths)
            {
                IMeasure measure = MeasureImporter.Load(path);
                measures.Add(measure);
            }

            return measures;
        }

        public Workbook LoadTemplateFile(string path)
        {
            Application app = new Application();
            Workbook workbook = new Workbook();

            app.Visible = false;
            workbook = app.Workbooks.Open(path);

            return workbook;
        }

        public LogMessage SaveMeasureFilesInExcel(IMeasureRepository repo, Workbook template)
        {
            LogMessage log = new LogMessage();

            string directoryToSave = $"{ repo.GetMeasuresDirectory() }\\Excel\\";
            Directory.CreateDirectory(directoryToSave);

            foreach (var measure in repo.Measures)
            {
                string measureName = measure.GetName();
                List<double> frictionFactors = measure.GetFrictionFactors();

                log.Info += $"{measureName} loaded.\n";

                CopyFrictionFactorsIntoExcel(frictionFactors, template);
                FormatChart(frictionFactors.Count, template);

                template.SaveCopyAs($"{directoryToSave}{measureName}.xlsx");

                log.Info += $"{measureName} saved.\n";
            }

            return log;
        }

        private void CopyFrictionFactorsIntoExcel(List<double> frictionFactors, Workbook template)
        {
            int currentRow = 1;
            Worksheet sheet = template.Worksheets[1];

            foreach (var f in frictionFactors)
            {
                sheet.Cells[currentRow++, "B"].Value = f;
            }
        }

        private void FormatChart(int NumberOfFrictionFactors, Workbook template)
        {
            double axisLenght = (NumberOfFrictionFactors + 2) / 10;

            Chart chart = (Chart)template.Charts[1];
            chart.Axes(XlAxisType.xlCategory).MaximumScale = axisLenght;
        }
    }
}

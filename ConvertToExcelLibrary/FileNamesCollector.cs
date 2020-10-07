using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertToExcelLibrary
{
    public static class FileNamesCollector
    {
        public static List<string> GetMeasureFromDialog()
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Multiselect = true;
            dlg.Filter = "Pliki *.MEASURE | *.MEASURE| Wszystkie pliki | *.*";
            dlg.ShowDialog();

            return dlg.FileNames.ToList();
        }

        public static string GetTemplateFromDialog()
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Pliki *.xls | *.xls| Wszystkie pliki | *.*";
            dlg.ShowDialog();

            return dlg.FileName;
        }
    }
}

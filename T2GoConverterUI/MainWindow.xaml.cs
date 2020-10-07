using ConvertToExcelLibrary;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
using Application = Microsoft.Office.Interop.Excel.Application;
using Window = System.Windows.Window;

namespace T2GoConverterUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        IMeasureRepository _repository = new MeasureRepository();
        IAppProcessor _app = new AppProcessor();
        string _templatePath = $"{Environment.CurrentDirectory}\\template.xlsx";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void GetMeasureFilesButton_Click(object sender, RoutedEventArgs e)
        {
            List<string> fileNames = FileNamesCollector.GetMeasureFromDialog();

            try
            {
                _repository.Measures = _app.LoadMeasureFiles(fileNames);
                ConvertToExcelButton.IsEnabled = true;
            }
            catch (XmlException ex)
            {
                MessageBox.Show($"Nieprawidłowy plik MEASURE lub plik jest uszkodzony\n\n{ ex.Message }", "Error");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Wystąpił nieznany błąd\n\n{ ex.Message }", "Error");
            }
        }

        private void ConvertToExcelButton_Click(object sender, RoutedEventArgs e)
        {
            bool conversionFinished = false;
            do
            {
                try
                {
                    ConvertMeasureToExcel();
                    conversionFinished = true;
                }
                catch (COMException ex)
                {
                    MessageBox.Show($"Plik szablonu nie został poprawnie załadowany\n\n{ ex.Message }", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    MessageBoxResult result = MessageBox.Show(
                        "Czy chcesz wczytać zewnętrzny plik szablonu?",
                        "Wczytaj szablon",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);

                    if (result.ToString() == "Yes")
                    {
                        _templatePath = FileNamesCollector.GetTemplateFromDialog();
                    }
                    else
                    {
                        MessageBox.Show("Applikacja zostanie zamknięta.", "Zamkanie aplikacji...", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Wystąpił błąd\n{ ex.Message }\n\nApplikacja zostanie zamknięta", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    conversionFinished = true;
                }
            } while (conversionFinished);
        }

        private void ConvertMeasureToExcel()
        {
            Workbook template = _app.LoadTemplateFile(_templatePath);
            try
            {
                LogMessage message = _app.SaveMeasureFilesInExcel(_repository, template);
                InfoTextBox.Text += "Konwersję zakończono pomyślenie.\n";
                InfoTextBox.Text += message.Info;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Pliki nie zostały przekonwertowane.\nZamknij program i spróbuj ponownie.\n\n{ ex.Message }", "Error");
                InfoTextBox.Text += "Wystąpił błąd. Zamknij program i spróbuj ponownie.";
                ConvertToExcelButton.IsEnabled = false;
                GetMeasureButton.IsEnabled = false;
            }
            finally
            {
                template.Close(SaveChanges: false);
            }
        }
    }
}

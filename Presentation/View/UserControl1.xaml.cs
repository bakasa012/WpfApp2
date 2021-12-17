using Ganss.Excel;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
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

namespace Presentation.View
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }
        string pathFile = "";
        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            getDataExcellFile("");
        }
        private void getDataExcellFile(string pathFile)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            List<DataBinddingExcel> dataBinddingExcels = new List<DataBinddingExcel>();
            try
            {
                string startEnd = "";
                var file = new System.IO.FileInfo(@"C:\Users\LIEU HONG THAI\source\repos\ExportImportExcelFile\ExportImportExcelFile\bin\Debug\Book1.xlsx");
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    if (package.Workbook.Worksheets.Count.ToString() != "0")
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        startEnd = worksheet.Dimension.Start.Row.ToString() + worksheet.Dimension.End.Row.ToString();
                        for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row + 1; i++)
                        {
                            try
                            {
                                int j = 1;
                                string name = worksheet.Cells[i, j++].Value?.ToString();
                                string code = worksheet.Cells[i, j++].Value?.ToString();
                                DataBinddingExcel dataBinddingExcel = new DataBinddingExcel()
                                {
                                    column1 = name,
                                    column2 = code
                                };
                                dataBinddingExcels.Add(dataBinddingExcel);
                            }
                            catch (Exception ex)
                            {

                                throw;
                            }
                        }

                    }
                    dtgExcel.ItemsSource = dataBinddingExcels;
                    textBlock1.Text = package.Workbook.Worksheets.Count.ToString() + " " + startEnd + " " + dataBinddingExcels.ToString();
                };
            }
            catch (Exception ex)
            {

                throw;
            }
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            ReadFileExceelWithExcelMapper("");
        }

        private void OpenFileDialog()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = ".xls";
            openFileDialog.Filter = "Text documents (.xls)|*.*";
            Nullable<bool> results = openFileDialog.ShowDialog();

            if (results == true)
            {
                FileNameTextBox.Text = openFileDialog.FileName;
                //textBlock1.Text = System.IO.File.ReadAllText(openFileDialog.FileName);
                pathFile = openFileDialog.FileName;
            }
        }

        private void ReadFileExceelWithExcelMapper(string pathFile)
        {
            var excel = new ExcelMapper(@pathFile);
            var products = excel.Fetch().ToList();
            textBlock1.Text = products.ToString();
            foreach (var item in products)
            {
                Console.WriteLine(item);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            ReadFileExceelWithExcelMapper(pathFile);
        }

        private void dtgExcel_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}

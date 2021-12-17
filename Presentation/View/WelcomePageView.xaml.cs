using Microsoft.Win32;
using OfficeOpenXml;
using Presentation.ViewModels;
using System;
using System.Windows;
using System.Windows.Controls;
using System.IO;
using Prism.Services.Dialogs;
using Ganss.Excel;
using System.Linq;

namespace Presentation.View
{
    /// <summary>
    /// Interaction logic for WelcomePageView.xaml
    /// </summary>
    public partial class WelcomePageView : UserControl
    {
        public WelcomePageView()
        {
            InitializeComponent();
            this.DataContext = new WelcomePageViewModel();
        }
        string pathFile = "";
        private void DialogResultButton(object sender, RoutedEventArgs e)
        {
            MessageBoxResult messageBoxResult = MessageBox.Show("confirm!", "Some title", MessageBoxButton.OKCancel);
            
            if (messageBoxResult == MessageBoxResult.OK) {
                textBlock1.Text = "true";
            }
            else if (messageBoxResult == MessageBoxResult.Cancel)
            {
                textBlock1.Text = "false1";
            }
            else
            {
                textBlock1.Text = "false2";
            }
        }
        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = ".xls";
            openFileDialog.Filter = "Text documents (.xls)|*.xls";
            Nullable<bool> results = openFileDialog.ShowDialog();

            if (results == true)
            {
                FileNameTextBox.Text = openFileDialog.FileName;
                //textBlock1.Text = System.IO.File.ReadAllText(openFileDialog.FileName);
                pathFile = openFileDialog.FileName;
            }
            
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(@"C:\Users\lieu.hong.thai\source\repos\WpfApp2\Presentation\bin\Debug\123.xlsx")))//new FileInfo("ImportData.xlsx")
            {
                textBlock1.Text = package.Workbook.Worksheets.Count.ToString();

                //textBlock1.Text = package.Workbook.Worksheets.Count.ToString();
            }


            /*List<DataBinddingExcel> dataBinddingExcels = new List<DataBinddingExcel>();
            try 
            {
                var package = new ExcelPackage( new System.IO.FileInfo("ImportData.xlsx"));
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                var test = package.Workbook.Worksheets;
                for (int i = worksheet.Dimension.Start.Row+1; i <= worksheet.Dimension.End.Row; i++)
                {
                    try
                    {
                        int j = 1;
                        string name = worksheet.Cells[i, j++].Value.ToString();
                        string code = worksheet.Cells[i, j++].Value.ToString();

                        DataBinddingExcel dataBinddingExcel = new DataBinddingExcel()
                        {
                            companyName = name,
                            companycode = code,
                        };
                        dataBinddingExcels.Add(dataBinddingExcel);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error : " + ex.ToString());
                        textBlock1.Text = ex.ToString();
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error : " + ex.ToString());
                textBlock1.Text = ex.ToString();

            }
            dtgExcel.ItemsSource = dataBinddingExcels;
        */
        }

        void ReadFileExcelWithExcelMapper(string pathFile)
        {
            var excel = new ExcelMapper(@pathFile);
            var dataBinddingExcel = excel.Fetch<DataBinddingExcel>().ToList();
            foreach (var item in dataBinddingExcel)
            {
                Console.WriteLine(item);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            ReadFileExcelWithExcelMapper(pathFile);
        }
    }
}

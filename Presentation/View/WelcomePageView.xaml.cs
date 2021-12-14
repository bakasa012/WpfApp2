using Microsoft.Win32;
using OfficeOpenXml;
using Presentation.ViewModels;
using System;
using System.Windows;
using System.Windows.Controls;
using System.IO;
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
            using (var package = new ExcelPackage(new FileInfo("ImportData.xlsx")))
            {
                //textBlock1.Text = package.Workbook.Worksheets.Count.ToString();
            }

            var package2 = new ExcelPackage(File.OpenRead(pathFile));

            textBlock1.Text = package2.Workbook.Worksheets.Count.ToString();
            int counter = 0;
            textBlock1.Text = "";
            foreach (string item in File.ReadLines(pathFile))
            {
                System.Console.WriteLine(item);
                textBlock1.Text += item.ToString()+counter.ToString();
                counter++;
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

        void ReadFileExcel()
        {
            //var package = new ExcelPackage(new System.IO.FileInfo("Book 1.xlsx"));
        }
    }
}

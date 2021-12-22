using Ganss.Excel;
using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.OpenXml4Net.OPC;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WpfApp2
{
    /// <summary>
    /// Interaction logic for Shell.xaml
    /// </summary>
    public partial class Shell : Window
    {
        private string pathFile;
        private string userLocal1 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)+ @"\DataExcelImport";
        private string pathFolderImportFileExcelGlobal = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\DataExcelImport";//@"C:\Users\lieu.hong.thai\Downloads\dataExcel";//@"C:\DataExcelImport";
        private string pathFolderExportFileExcelGlobal = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\DataExcelExport";//@"C:\DataExcelExport\";
        //int rowIndex = 0;
        private List<DataBinddingExcel> dataBinddingExcels = new List<DataBinddingExcel>();
        private List<DataBodyExcelFile> dataBodyExcelFiles = new List<DataBodyExcelFile>();
        List<string> files = new List<string>();

        public Shell()
        {
            InitializeComponent();
            checkPathOrCreateFolder(pathFolderImportFileExcelGlobal);
            checkPathOrCreateFolder(pathFolderExportFileExcelGlobal);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog();

        }
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            processConverFileExcel();
        }
        private void OpenFileDialog()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = ".xls";
            openFileDialog.Filter = "Text documents (.xls)|*.*";
            Nullable<bool> results = openFileDialog.ShowDialog();

            if (results == true)
            {
                pathFile = openFileDialog.FileName;
                ReadFileExcelWitdNPOI(openFileDialog.FileName);
            }
        }
        private void checkPathOrCreateFolder(string path)
        {
            if (Directory.Exists(path))
            {

            }
            else
            {
                Directory.CreateDirectory(path);
            }
        }

        private void processConverFileExcel()
        {
            //C:\DataExcelImport\3870_GEO北九州三ヶ森店-【10月末〆】10月度自己調達許諾シール給付申請書 (2).xls
            recursiveDirectory(pathFolderImportFileExcelGlobal);
            string[] files = this.files.ToArray();
            //string[] files = Directory.GetFiles(pathFolderImportFileExcelGlobal);
            for (int i = 0; i < files.Length; i++)
            {
                string fileExt = System.IO.Path.GetExtension(files[i]);
                string[] folder = System.IO.Directory.GetDirectories(pathFolderImportFileExcelGlobal);
                if (fileExt == ".xls" || fileExt == ".xlsx")
                {
                    textBlock1.Text = files[i];
                    this.ReadFileExcelWitdNPOI(@files[i]);
                    this.ExportFileExcel(pathFolderExportFileExcelGlobal + @"\" + System.IO.Path.GetFileName(files[i]).ToString().Split('.')[0] + i.ToString() + ".xlsx");//("./demo.xlsx");
                }

            }
            this.files.Clear();
            string strPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            textBlock1.Text = strPath;
            MessageBox.Show(pathFolderImportFileExcelGlobal + " " + "Files was conver : " + files.Length.ToString(), "Information!"+ userLocal1);

        }
        private void recursiveDirectory(string path)
        {
            string[] folder = Directory.GetDirectories(path);
            for (int i = 0; i < Directory.GetFiles(path).Length; i++)
            {
                files.Add(Directory.GetFiles(path)[i]);

            }
            if (folder.Length >0)
            {
                for (int i = 0; i < folder.Length; i++)
                {
                    recursiveDirectory(folder[i]);
                }
            }
            //else
            //{

            //    for (int i = 0; i < Directory.GetFiles(path).Length; i++)
            //    {
            //        files.Add(Directory.GetFiles(path)[i]);

            //    }
            //}
        }
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            ReadFileExcelWitdNPOI(pathFile);
        }
        private void ReadFileExceelWithExcelMapper(string pathFile)
        {
            var excel = new ExcelMapper(@pathFile);
            var products = excel.Fetch<DataBinddingExcel>().ToList();
            foreach (var item in products)
            {
                Console.WriteLine(item);
            }
        }
        private void ReadFileExcelWitdNPOI(string pathFile)
        {
            
            //OPCPackage pkg = OPCPackage.Open(pathFile);
            IWorkbook wb = null;
            dataBinddingExcels.Clear();
            dataBodyExcelFiles.Clear();
            using (FileStream fileStream = new FileStream(@pathFile, FileMode.Open, FileAccess.Read)) 
            {
                string fileExt = System.IO.Path.GetExtension(pathFile);
                try
                {
                    switch (fileExt.ToLower())
                    {
                        case ".xls":
                            wb = new HSSFWorkbook(fileStream);
                            break;
                        case ".xlsx":
                            wb = new XSSFWorkbook(fileStream);
                            break;
                        default:
                            break;
                    }
                }
                catch (Exception ex)
                {
                    //wb = new XSSFWorkbook(fileStream);
                }
                /*finally
                {
                    wb = new HSSFWorkbook(fileStream);
                }*/

                ISheet sheet = wb.GetSheetAt(0);
                int lastRow = sheet.LastRowNum;
                bool isDataBody = false;
                int rowIndex = 0;
                while (true)
                {
                    var nowRow = sheet.GetRow(rowIndex);
                    if (nowRow != null)
                    {
                        if (nowRow.GetCell(1)?.ToString() == "ISBNコード")
                        {
                            isDataBody = true;
                        }
                        if (!isDataBody)
                        {
                            DataBinddingExcel dataBinddingExcel = new DataBinddingExcel()
                            {
                                column1 = nowRow.GetCell(0)?.ToString(),
                                column2 = nowRow.GetCell(1)?.ToString(),
                                column3 = nowRow.GetCell(2)?.ToString().Trim(),
                                column4 = nowRow.GetCell(3)?.ToString(),
                                column5 = nowRow.GetCell(4)?.ToString(),
                                column6 = nowRow.GetCell(5)?.ToString(),
                                column7 = nowRow.GetCell(6)?.ToString(),
                                column8 = nowRow.GetCell(7)?.ToString(),
                            };
                            dataBinddingExcels.Add(dataBinddingExcel);
                        }
                        else
                        {
                            DataBodyExcelFile dataBodyExcelFile = new DataBodyExcelFile()
                            {
                                column1 = nowRow.GetCell(0)?.ToString(),
                                column2 = nowRow.GetCell(1)?.ToString(),
                                column3 = nowRow.GetCell(2)?.ToString(),
                                column4 = nowRow.GetCell(3)?.ToString(),
                                column5 = nowRow.GetCell(4)?.ToString(),
                                column6 = nowRow.GetCell(5)?.ToString(),
                                column7 = nowRow.GetCell(6)?.ToString(),
                                column8 = nowRow.GetCell(7)?.ToString(),
                            };
                            dataBodyExcelFiles.Add(dataBodyExcelFile);
                        }

                    }
                    else if(!isDataBody)
                    {
                        DataBinddingExcel dataBinddingExcel = new DataBinddingExcel()
                        {
                            column1 = "",
                            column2 = "",
                            column3 = "",
                            column4 = "",
                            column5 = "",
                            column6 = "",
                            column7 = "",
                            column8 = "",
                        };
                        dataBinddingExcels.Add(dataBinddingExcel);
                    }
                    if (rowIndex >= lastRow)
                        break;
                    rowIndex++;
                }

                dtgExcelReport2.ItemsSource = dataBinddingExcels;
                fileStream.Close();
            };
        }
        private void ExportFileExcel(string output)
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            var row = sheet.CreateRow(0);
            row.CreateCell(1);
            //Merge column
            CellRangeAddress cellRange = new CellRangeAddress(0, 0, 1, 7);
            sheet.AddMergedRegion(cellRange);
            sheet.AddMergedRegion(new CellRangeAddress(2, 2, 5, 7));
            sheet.AddMergedRegion(new CellRangeAddress(3, 3, 5, 7));
            sheet.AddMergedRegion(new CellRangeAddress(4, 4, 3, 7));
            sheet.AddMergedRegion(new CellRangeAddress(5, 5, 3, 7));
            sheet.AddMergedRegion(new CellRangeAddress(6, 6, 3, 4));
            sheet.AddMergedRegion(new CellRangeAddress(6, 6, 5, 7));
            sheet.AddMergedRegion(new CellRangeAddress(7, 7, 4, 7));
            //set column width
            sheet.AutoSizeColumn(0);
            sheet.SetColumnWidth(1, 4000);
            sheet.SetColumnWidth(2, 6000);
            sheet.SetColumnWidth(3, 4000);
            sheet.SetColumnWidth(4, 3000);
            sheet.SetColumnWidth(5, 3000);
            sheet.SetColumnWidth(6, 3000);
            sheet.SetColumnWidth(7, 3000);
            sheet.SetColumnWidth(8, 6000);
            //end

            row.GetCell(1).SetCellValue(dataBinddingExcels[0].column2);
            FontChange(wb, "title", row, 1);

            int rowIndex = 1;
            dataBinddingExcels.RemoveAt(0);
            dataBinddingExcels.RemoveAt(dataBinddingExcels.Count - 1);
            //data header
            foreach (var item in dataBinddingExcels)
            {
                var newRow = sheet.CreateRow(rowIndex);
                newRow.CreateCell(0).SetCellValue(item.column1);
                newRow.CreateCell(1).SetCellValue(item.column2);
                
                newRow.CreateCell(2).SetCellValue(item.column3);
                
                newRow.CreateCell(3).SetCellValue(item.column4);

                newRow.CreateCell(4).SetCellValue(item.column5);
                
                newRow.CreateCell(5).SetCellValue(item.column6);
                
                newRow.CreateCell(6).SetCellValue(item.column7);
                newRow.CreateCell(7).SetCellValue(item.column8);
                if (rowIndex < 8)
                    for (int i = 1; i <= 7; i++)
                    {
                        FontChange(wb, "content", newRow, i);
                    }
                else
                    for (int i = 1; i <= 7; i++)
                        FontChange(wb, "description",newRow, i);
                if (item.column2 != null && item.column2 != "" && item.column2 != "null" && rowIndex <= 8 )
                {
                    FontChange(wb, "label", newRow, 1);
                }
                if (rowIndex == 3)
                {
                    FontChange(wb, "bgBlack", newRow, 5);
                    FontChange(wb, "bgBlack", newRow, 6);
                    FontChange(wb, "bgBlack", newRow, 7);
                }
                if (rowIndex == 4)
                {
                    FontChange(wb, "bgBlack", newRow, 3);
                    FontChange(wb, "bgBlack", newRow, 4);
                    FontChange(wb, "bgBlack", newRow, 5);
                    FontChange(wb, "bgBlack", newRow, 6);
                    FontChange(wb, "bgBlack", newRow, 7);
                }
                rowIndex++;
            }

            //data body
            rowIndex++;
            foreach(var item in dataBodyExcelFiles)
            {
                var newRow = sheet.CreateRow(rowIndex);
                newRow.CreateCell(0).SetCellValue(item.column1);
                newRow.CreateCell(1).SetCellValue(item.column2);

                newRow.CreateCell(2).SetCellValue(item.column3);
                newRow.CreateCell(3).SetCellValue(item.column4);
                newRow.CreateCell(4).SetCellValue(item.column5);
                newRow.CreateCell(5).SetCellValue(item.column6);
                newRow.CreateCell(6).SetCellValue(item.column7);
                newRow.CreateCell(7).SetCellValue(item.column8);
                for (int i = 1; i <= 7; i++)
                {
                    FontChange(wb, "content", newRow, i);
                }
                rowIndex++;
            }




            if (File.Exists(output))
                File.Delete(output);
            FileStream fileStream = new FileStream(output, FileMode.CreateNew,FileAccess.ReadWrite,FileShare.None);
            wb.Write(fileStream);
            fileStream.Close();
        }

        private void FontChange(XSSFWorkbook wb, string caseFont, IRow row, int index)
        {
            IFont font = wb.CreateFont();
            ICellStyle cellStyle = wb.CreateCellStyle();

            switch (caseFont)
            {
                case "title":
                    font.Boldweight = (short)FontBoldWeight.Bold;
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                    font.FontName = "MS PGothic";
                    font.FontHeightInPoints = 16;
                    break;
                case "label":
                    font.FontName = "MS PGothic";
                    font.FontHeightInPoints = 11;
                    cellStyle.FillForegroundColor = HSSFColor.Gold.Index;
                    break;
                case "content":
                    font.FontName = "MS PGothic";
                    font.FontHeightInPoints = 11;
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                    /*cellStyle.BorderBottom = BorderStyle.Thin;
                    cellStyle.BorderLeft = BorderStyle.Thin;
                    cellStyle.BorderRight = BorderStyle.Thin;
                    cellStyle.BorderTop = BorderStyle.Thin;*/
                    break;
                case "description":
                    font.FontName = "MS PGothic";
                    font.FontHeightInPoints = 11;
                    break;
                case "bgBlack":
                    /*cellStyle.BorderBottom = BorderStyle.None;
                    cellStyle.BorderLeft = BorderStyle.None;
                    cellStyle.BorderRight = BorderStyle.None;
                    cellStyle.BorderTop = BorderStyle.None;*/
                    cellStyle.FillForegroundColor = HSSFColor.Grey50Percent.Index;
                    cellStyle.FillBackgroundColor = HSSFColor.Red.Index;
                    break;
                default:
                    font.FontName = "Calibri";
                    cellStyle.FillBackgroundColor = HSSFColor.Gold.Index;
                    break;
            }
            cellStyle.SetFont(font);
            row.GetCell(index).CellStyle = cellStyle;
        }

    }
}

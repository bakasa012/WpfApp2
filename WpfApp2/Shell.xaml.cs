using Ganss.Excel;
using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
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

        public Shell()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog();

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
        int rowIndex = 0;
        List<DataBinddingExcel> dataBinddingExcels = new List<DataBinddingExcel>();
        List<DataBodyExcelFile> dataBodyExcelFiles = new List<DataBodyExcelFile>();
        private void ReadFileExcelWitdNPOI(string pathFile)
        {
            dataBinddingExcels.Clear();
            dataBodyExcelFiles.Clear();
            FileStream fileStream = new FileStream(@pathFile, FileMode.Open);
            HSSFWorkbook wb = new HSSFWorkbook(fileStream);
            ISheet sheet = wb.GetSheetAt(0);
            int lastRow = sheet.LastRowNum;
            bool isDataBody = false;
            while (true)
            {
                var nowRow = sheet.GetRow(rowIndex);
                if (nowRow != null)
                {
                    var MSSV = nowRow.GetCell(0)?.StringCellValue;
                    var Name = nowRow.GetCell(1)?.ToString();
                    var Phone = nowRow.GetCell(2)?.StringCellValue;
                    if(nowRow.GetCell(1)?.ToString() == "ISBNコード")
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
                if (rowIndex >= lastRow)
                    break;
                rowIndex++;
            }

            dtgExcelReport2.ItemsSource = dataBinddingExcels;
            fileStream.Close();
        }
        private void ExportFileExcel(string output)
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            var row = sheet.CreateRow(0);
            row.CreateCell(1);
            //background color
            //cellStyleLabel.FillForegroundColor = HSSFColor.Gold.Index;
            //end
            //Merge column
            CellRangeAddress cellRange = new CellRangeAddress(0,0,1,6);
            sheet.AddMergedRegion(cellRange);
            sheet.AddMergedRegion(new CellRangeAddress(3, 3, 5, 6));
            sheet.AddMergedRegion(new CellRangeAddress(4, 4, 5, 6));
            sheet.AddMergedRegion(new CellRangeAddress(5, 5, 3, 6));
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
                if (rowIndex != 8 && rowIndex != 9)
                for (int i = 1; i <= 7; i++)
                {
                    FontChange(wb, "content", newRow, i);
                }
                if (item.column2 != null && item.column2 != "" && item.column2 != "null" && rowIndex != 8 && rowIndex != 9)
                {
                    FontChange(wb, "label", newRow, 1);
                }
                if (rowIndex == 4)
                {
                    FontChange(wb, "bgBlack", newRow, 3);
                    FontChange(wb, "bgBlack", newRow, 4);
                    FontChange(wb, "bgBlack", newRow, 5);
                    FontChange(wb, "bgBlack", newRow, 6);
                }
                if (rowIndex == 3)
                {
                    FontChange(wb, "bgBlack", newRow, 5);
                    FontChange(wb, "bgBlack", newRow, 6);
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
                rowIndex++;
            }




            if (File.Exists(output))
                File.Delete(output);
            FileStream fileStream = new FileStream("./demo.xlsx", FileMode.CreateNew,FileAccess.ReadWrite,FileShare.None);
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


        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            ExportFileExcel("./demo.xlsx");
        }
    }
}

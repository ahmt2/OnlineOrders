using CsvHelper.Configuration;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace OnlineOrders
{
    public partial class OrdersRibbon
    {
        readonly Random rnd = new Random();
        Excel.Workbook wBook;
        Excel.Worksheet wSheet;
        string firstFile;
        string pathForFiles;
        string orderRange;
        bool csvFilesLoaded;
        private SortedList<Product, uint> sortedPList = new SortedList<Product, uint>();

        private void InitializeOpenFileDialog()
        {
            this.openFileDialog1.Filter = "CSV Files (*.csv)|*.csv";

            // Allow the user to select multiple files.
            this.openFileDialog1.Multiselect = true;
            this.openFileDialog1.Title = "Select order file(s)";
        }

        private void OrdersRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            InitializeOpenFileDialog();
            csvFilesLoaded = false;
        }

        /// <summary>
        /// Find order numbers (first and last) from the file list
        /// Order file name format is: YYMMDDsometext1234 so if it's in 
        /// that format we need to get the order number. If not, no need 
        /// for an order number. 
        /// </summary>
        /// <param name="filePaths"></param>
        /// <returns></returns>
        private string GetOrderRange(List<string> filePaths)
        {
            string range = "";
            string fileName;
            filePaths.Sort();
            var regex = new Regex(@"[0-9]{6}.+[0-9]{4}");
            
            //get the file name by splitting using \\, then get numbers and drop the first 6 since they are for date
            int DATE_STR_LENGTH = 6;
            string firstFileName = filePaths[0].Split('\\')[filePaths[0].Split('\\').Length - 1].Split('.')[0];
            bool nonMatchingNameExists = false;
            foreach(String filePath in filePaths)
            {
                fileName = filePath.Split('\\')[filePath.Split('\\').Length - 1].Split('.')[0];
                var results = regex.Matches(fileName);
                if (0 == results.Count)
                {
                    nonMatchingNameExists = true;
                }
            }
            
            if(nonMatchingNameExists)
            {
                return range;
            }
            string firstFileIndex = string.Join("", firstFileName.ToCharArray().Where(Char.IsDigit)).Substring(DATE_STR_LENGTH);
            range = "Order " + firstFileIndex;
            if(filePaths.Count > 1)
            {                
                string lastPath = filePaths[filePaths.Count - 1].Split('\\')[filePaths[filePaths.Count - 1].Split('\\').Length - 1].Split('.')[0];
                string lastFileIndex = string.Join("", lastPath.ToCharArray().Where(Char.IsDigit)).Substring(DATE_STR_LENGTH);
                range += ("-" + lastFileIndex);
            }
            return range;
        }

        private void BtnLoadCSVFiles_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp =
                (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

            System.Windows.Forms.DialogResult dr = this.openFileDialog1.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                pathForFiles = Path.GetDirectoryName(openFileDialog1.FileNames[0]);
                firstFile = openFileDialog1.FileNames[0];
                wSheet = (Excel.Worksheet)excelApp.ActiveWorkbook.ActiveSheet;
                double noOfUsedCells = excelApp.WorksheetFunction.CountA(wSheet.Cells);
                if (noOfUsedCells > 0)
                {   //current worksheet not empty, switch to a new worksheet
                    Excel.Worksheet newWorksheet;
                    newWorksheet = (Excel.Worksheet)excelApp.ActiveWorkbook.Worksheets.Add();
                    try
                    {
                        newWorksheet.Name = firstFile.Split('\\')[firstFile.Split('\\').Length - 1].Split('.')[0];
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {   //worksheet name already exists, append a random number for now 
                        int rNum = rnd.Next(1000, 10000);    
                        newWorksheet.Name = firstFile.Split('\\')[firstFile.Split('\\').Length - 1].Split('.')[0] + rNum;
                    }
                }
                wSheet = (Excel.Worksheet)excelApp.ActiveWorkbook.ActiveSheet;
                wBook = excelApp.ActiveWorkbook;
                wSheet.Cells.NumberFormat = "General";

                var reader = new StreamReader(openFileDialog1.FileNames[0]);
                var config = new CsvConfiguration(CultureInfo.InvariantCulture) { Delimiter = ",", Encoding = Encoding.UTF8 };
                config.BadDataFound = null;
                var csv = new CsvHelper.CsvReader(reader, config);
                CsvHelper.CsvDataReader csvDataReader = new CsvHelper.CsvDataReader(csv);
                //Rows.Results View[0].ItemArray[0] : BC3001 
                //Columns.List[0] : Product
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(csvDataReader);
                //uint rowCount = ((uint)dt.Rows.Count);
                var Datarow = dt.Select("");
                //System.Diagnostics.Debug.WriteLine(Datarow[0].ItemArray[2]);

                for (int i = 0; i < dt.Columns.Count; i++)
                {//first cell array index: row, second index: column
                    wSheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
                }
                wSheet.get_Range("A1", "E1").Font.Bold = true;
                reader.Close();

                Product temp;
                orderRange = GetOrderRange(new List<string>(openFileDialog1.FileNames));

                foreach (String file in openFileDialog1.FileNames)
                {
                    reader = new StreamReader(file);
                    csv = new CsvHelper.CsvReader(reader, config);
                    csvDataReader = new CsvHelper.CsvDataReader(csv);
                    dt = new System.Data.DataTable();
                    dt.Load(csvDataReader);
                    Datarow = dt.Select("");

                    for (uint i = 0; i < Datarow.Length; i++)
                    {
                        temp = new Product(Datarow[i].ItemArray);
                        uint quantity = uint.Parse((string)Datarow[i].ItemArray[4]);

                        if (sortedPList.ContainsKey(temp))
                        {
                            sortedPList[temp] += quantity;
                        }
                        else
                        {
                            sortedPList.Add(temp, quantity);
                        }
                    }
                    reader.Close();
                }

                int currentRow = 2;
                string currentStyle = sortedPList.Keys[0].Style;
                foreach (KeyValuePair<Product, uint> keyValue in sortedPList)
                {
                    if(keyValue.Key.Style.CompareTo(currentStyle) != 0)
                    {
                        currentRow += 2;
                        currentStyle = keyValue.Key.Style;
                    }
                    
                    wSheet.Cells[currentRow, 1] = keyValue.Key.Style;
                    wSheet.Cells[currentRow, 2] = keyValue.Key.Color;
                    wSheet.Cells[currentRow, 3] = "'" + keyValue.Key.ProductCode;
                    wSheet.Cells[currentRow, 4] = keyValue.Key.Size;
                    wSheet.Cells[currentRow, 5] = keyValue.Value;
                    currentRow += 1;
                }
                csvFilesLoaded = true;
            }
        }

        private void BtnGenerate_Click(object sender, RibbonControlEventArgs e)
        {
            if(false == csvFilesLoaded)
            {
                System.Windows.Forms.MessageBox.Show(
                    "You should first load a .csv file", "Error", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return;
            }
            var applicationWord = new Word.Application
            {
                Visible = true
            };
            Word.Document wordDoc = applicationWord.Documents.Add();
            Range xlRange = wSheet.UsedRange;
            xlRange.Copy();
            wordDoc.ActiveWindow.Selection.PasteExcelTable(false, true, false);

            int noOfTables = wordDoc.Tables.Count;
            if(noOfTables > 0)
            {   //Tables collection is null at index 0 
                Word.Table wordTable = wordDoc.Tables[1];
                wordTable.AllowAutoFit = true;
                wordTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
            }

            DateTime today = DateTime.Today;
            object missing = System.Reflection.Missing.Value;
            foreach (Word.Section section in wordDoc.Sections)
            {
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                headerRange.Font.Bold = 1;
                headerRange.Text = orderRange + "\tHOUSTON STORE\t" + today.ToString("MM.dd.yyyy");
            }
            //System.Diagnostics.Debug.WriteLine(ddbPartSelector.SelectedItem.ToString());
            //https://www.programmersought.com/article/30982190787/
            object CurrentPage = Word.WdFieldType.wdFieldPage;
            object TotalPages = Word.WdFieldType.wdFieldNumPages;
            //open footer
            applicationWord.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
            applicationWord.ActiveWindow.Selection.TypeText("Part " + ddbPartSelector.SelectedItem.ToString() + "\tPage ");
            applicationWord.ActiveWindow.Selection.Fields.Add(applicationWord.ActiveWindow.Selection.Range, ref CurrentPage, ref missing, ref missing);
            applicationWord.ActiveWindow.Selection.TypeText(" of ");
            applicationWord.ActiveWindow.Selection.Fields.Add(applicationWord.ActiveWindow.Selection.Range, ref TotalPages, ref missing, ref missing);
            //close footer
            applicationWord.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

            object filename = @pathForFiles + "\\HoustonStore-Part" + ddbPartSelector.SelectedItem.ToString() + "-" + today.ToString("MM.dd.yyyy") + ".docx";
            wordDoc.SaveAs2(ref filename);
            //wordDoc.Close(ref missing, ref missing, ref missing);
            //wordDoc = null;
            //applicationWord.Quit(ref missing, ref missing, ref missing);
            //applicationWord = null;
        }
    }
}

using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Globalization;
using CsvHelper.Configuration;

namespace OnlineOrders
{
    public partial class Ribbon1
    {
        readonly Random rnd = new Random();
        Excel.Workbook wBook;
        Excel.Worksheet wSheet;
        string firstFile;
        string pathForFiles;
        private SortedList<Product, uint> sortedPList = new SortedList<Product, uint>();

        private void InitializeOpenFileDialog()
        {
            this.openFileDialog1.Filter = "CSV Files (*.csv)|*.csv";

            // Allow the user to select multiple files.
            this.openFileDialog1.Multiselect = true;
            this.openFileDialog1.Title = "Select order file(s)";
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            InitializeOpenFileDialog();
        }

        private void BtnOnlineOrders_Click(object sender, RibbonControlEventArgs e)
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
            }
        }

        private void BtnGenerate_Click(object sender, RibbonControlEventArgs e)
        {
            var applicationWord = new Word.Application
            {
                Visible = true
            };
            Word.Document wordDoc = applicationWord.Documents.Add();
            Range xlRange = wSheet.UsedRange;
            xlRange.Copy();
            wordDoc.ActiveWindow.Selection.PasteExcelTable(false, true, false);
        }
    }
}

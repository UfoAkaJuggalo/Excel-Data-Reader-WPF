using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace IPG_zad2.Model
{
    public class ExcelDataModel
    {
        public FileDataModel fileDataModel;

        public ExcelDataModel() => fileDataModel = new FileDataModel();

        public void ReadFile (string filePath)
        {
            Application xlApp = new Application();
            Workbook xlWorkBook = xlApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            
            foreach (Worksheet sheet in xlWorkBook.Worksheets)
            {
                Range range = sheet.UsedRange;
                SheetModel modelSheet = new SheetModel();
                for (int currColumn = 1; currColumn < range.Columns.Count; currColumn++)
                {
                    string columnName = (string)(range.Cells[1, currColumn] as Range).Value2;
                    switch (columnName.Trim().ToLower())
                    {
                        case "id":
                            modelSheet.Id = GetColumnNumberValuesList(currColumn, range);
                            break;
                        case "nazwa":
                            modelSheet.Name = GetColumnStringValuesList(currColumn, range);
                            break;
                        case "cena":
                            modelSheet.Price = GetColumnPriceValuesList(currColumn, range);
                            break;
                        case "pozycja":
                            modelSheet.Position = GetColumnNumberValuesList(currColumn, range);
                            break;
                        case "poziom":
                            modelSheet.Level = GetColumnStringValuesList(currColumn, range);
                            break;
                        case "opis":
                            modelSheet.Description = GetColumnStringValuesList(currColumn, range);
                            break;
                        case "nr zamówienia":
                            modelSheet.Order = GetColumnStringValuesList(currColumn, range);
                            break;
                        default:
                            break;
                    }
                }
                fileDataModel.SheetList.Add(modelSheet);
            }
            xlWorkBook.Close(false,null,null);
            xlApp.Quit();
        }

        private List<string> GetColumnStringValuesList (int colNum, Range range)
        {
            List<string> retList = new List<string>();
            for (int currRow = 2; currRow < range.Rows.Count; currRow++)
            {
                retList.Add((string)(range.Cells[currRow, colNum] as Range).Value2);
            }
            return retList;
        }
        private List<int> GetColumnNumberValuesList(int colNum, Range range)
        {
            List<int> retList = new List<int>();
            for (int currRow = 2; currRow < range.Rows.Count; currRow++)
            {
                if ((range.Cells[currRow, colNum] as Range).Value2 != null)
                    retList.Add((int)(range.Cells[currRow, colNum] as Range).Value2);
                else
                    retList.Add(0);
            }
            return retList;
        }
        private List<int> GetColumnPriceValuesList(int colNum, Range range)
        {
            List<int> retList = new List<int>();
            for (int currRow = 2; currRow < range.Rows.Count; currRow++)
            {
                retList.Add((int)(range.Cells[currRow, colNum] as Range).Value2);
            }
            return retList;
        }
    }
}

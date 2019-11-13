using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
namespace EY.Business.ExcelOperations
{
    public class ExcelOperations
    {
        public int RowCount { get; set; }
        public List<string> GetAllHeaders(string FilePath, int SheetNo)
        {
            List<string> excelHeaders = new List<string>();

            //excelHeaders.Add("--Select--");
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(FilePath, false))
            {
                WorkbookPart wbPart = doc.WorkbookPart; // creates an instance of the workbook
                int worksheetcount = doc.WorkbookPart.Workbook.Sheets.Count(); //counts the total number of worksheets
                Sheet mysheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(SheetNo);
                Worksheet Worksheet = ((WorksheetPart)wbPart.GetPartById(mysheet.Id)).Worksheet;
                //SheetData Rows = (SheetData)Worksheet.ChildElements.GetItem(4);
                SheetData Rows = (SheetData)Worksheet.GetFirstChild<SheetData>();
                Row currentrow = (Row)Rows.ChildElements.GetItem(0);
                int columnCount = currentrow.ChildElements.Count;
                for (int i = 0; i < columnCount; i++)
                {
                    Cell currentcell = (Cell)currentrow.ChildElements.GetItem(i);
                    string currentcellvalue = string.Empty;

                    if (currentcell.DataType != null)
                    {
                        if (currentcell.DataType == CellValues.SharedString)
                        {
                            int id = -1;

                            if (Int32.TryParse(currentcell.InnerText, out id))
                            {
                                SharedStringItem item = GetSharedStringItemById(wbPart, id);

                                if (item.Text != null)
                                {
                                    //code to take the string value  
                                    currentcellvalue = item.Text.Text;
                                }
                                else
                                {
                                    currentcellvalue = "Empty Header";
                                }
                            }
                        }
                    }
                    excelHeaders.Add(currentcellvalue);


                }
                doc.Close();

            }

            return excelHeaders;
        }
        public static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id)
        {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }

        public static bool CheckValidExcelFile(string FilePath, string FileName)
        {
            string Extension = FileName.Split('.')[1];
            bool fileExists = File.Exists(FilePath);
            if (fileExists)
            {

                if (Extension == "xls" || Extension == "xlsx")
                {
                    try
                    {


                        using (SpreadsheetDocument doc = SpreadsheetDocument.Open(FilePath, false))
                        {
                            WorkbookPart wbPart = doc.WorkbookPart;
                            int worksheetcount = doc.WorkbookPart.Workbook.Sheets.Count();
                            if (worksheetcount == 1)
                                return true;
                            else
                                return false;
                        }
                    }
                    catch (Exception ex)
                    {

                        throw;
                    }
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// return all sheetnames in the uploaded file
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        public List<string> GetAllSheetNames(string filepath)
        {
            List<string> workSheets = new List<string>();

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filepath, false))
            {
                WorkbookPart wbPart = doc.WorkbookPart; // creates an instance of the workbook
                int worksheetcount = doc.WorkbookPart.Workbook.Sheets.Count(); //counts the total number of worksheets

                foreach (var item in doc.WorkbookPart.Workbook.Sheets.ChildElements)
                {
                    string sheetname = ((Sheet)item).Name;
                    workSheets.Add(sheetname);
                }
                doc.Close();
            }

            return workSheets;
        }

        /// <summary>
        /// return all column names
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        public List<string> GetAllColumnHeaders(string filepath, string sheetname)
        {
            List<string> excelHeaders = new List<string>();
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filepath, false))
            {
                WorkbookPart wbPart = doc.WorkbookPart; // creates an instance of the workbook
                                                        // int worksheetcount = doc.WorkbookPart.Workbook.Sheets.Count(); //counts the total number of worksheets
                foreach (var items in doc.WorkbookPart.Workbook.Sheets.ChildElements)
                {
                    Sheet mysheet = (Sheet)items;
                    if (mysheet.Name == sheetname)
                    {
                        Worksheet Worksheet = ((WorksheetPart)wbPart.GetPartById(mysheet.Id)).Worksheet;

                        SheetData Rows = (SheetData)Worksheet.GetFirstChild<SheetData>();
                        this.RowCount = Rows.Count();
                        //SheetData Rows = (SheetData)Worksheet.ChildElements.GetItem(4);
                        if (Rows.ChildElements.Count > 0)
                        {
                            Row currentrow = (Row)Rows.ChildElements.GetItem(0);
                            int columnCount = currentrow.ChildElements.Count;
                            for (int i = 0; i < columnCount; i++)
                            {
                                Cell currentcell = (Cell)currentrow.ChildElements.GetItem(i);
                                string currentcellvalue = string.Empty;

                                if (currentcell.DataType != null)
                                {
                                    if (currentcell.DataType == CellValues.SharedString)
                                    {
                                        int id = -1;

                                        if (Int32.TryParse(currentcell.InnerText, out id))
                                        {
                                            SharedStringItem item = GetSharedStringItemById(wbPart, id);

                                            if (item.Text != null)
                                            {
                                                //code to take the string value
                                                currentcellvalue = item.Text.Text.ToLower();
                                            }
                                            else
                                            {
                                                currentcellvalue = "Empty Header";
                                            }
                                        }
                                    }
                                }
                                excelHeaders.Add(currentcellvalue);
                            }
                        }
                    }

                }
                doc.Close();
            }
            return excelHeaders;
        }
    }

}

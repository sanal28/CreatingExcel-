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

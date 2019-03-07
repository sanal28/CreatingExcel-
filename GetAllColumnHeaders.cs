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

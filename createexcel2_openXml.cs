  public EmailParams CreateUploadExcelDocRS(DataSet dataSet, ParameterType pType, string code, string FilePath)
        {
           // string s = "C:\\temp\\p.xlsx";
            var retVal = new EmailParams(pType);
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(FilePath, SpreadsheetDocumentType.Workbook))
            {
                
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();
                


                for (int tblIdx = 0; tblIdx < dataSet.Tables.Count; tblIdx++)
                {
                    var table = dataSet.Tables[tblIdx];
                    Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                    Sheet sheet = new Sheet()
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Parameter File"
                    };

                    sheets.Append(sheet);

                    workbookPart.Workbook.Save();


                    SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                    // Constructing header

                    Row headerRow = new Row();

                    List<String> columns = new List<string>();
                    foreach (DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);

                        Cell cell = new Cell
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(column.ColumnName)
                            
                        };
                        headerRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(headerRow);

                    foreach (DataRow dsrow in table.Rows)
                    {
                        Row newRow = new Row();
                        foreach (String col in columns)
                        {
                            var cellText = dsrow[col].ToString().Trim();
                            Cell cell = new Cell
                            {
                                CellValue = new CellValue(cellText),
                                DataType = new EnumValue<CellValues>(CellValues.String)


                                //DataType = CellValues.String,
                                //CellValue = new CellValue(cellText)
                            };
                            newRow.AppendChild(cell);
                            if (tblIdx == 0) // only needed for Initial DataTable
                            {
                                FillEmailParams(retVal, col, cellText);
                            }
                        }

                        sheetData.AppendChild(newRow);
                    }
                }
                    worksheetPart.Worksheet.Save();
                
            }
           return retVal ;
        }

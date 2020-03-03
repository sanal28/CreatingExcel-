using OfficeOpenXml;

controler
 
 filename=sfsfsf;
  templateList = await _storedProcedure.GetGlobalAutomationTemplateData(officeID, userID);
 byte[] fileContents = _excelService.DATargetDownload(templateList, type, office);
 
  if (fileContents == null || fileContents.Length == 0)
            {
                return NotFound();
            }
            return File(fileContents, "application/octet-stream", filename + ".xlsx");
 service
 
   public byte[] DATargetDownload(DADownloadTemplate DATasks, string DAName, string Office)
        {
            using (var package = new ExcelPackage())
            {
                var sheetName = $"{DAName + "_" + Office + "_" + DateTime.Now.Date.ToShortDateString()}";
                var workSheet = package.Workbook.Worksheets.Add(sheetName);
                if (DAName == "GL")
                    workSheet.Cells.LoadFromCollection(DATasks.GLSLTemplateList, true);
                
                if (DAName == "GA")//Global Automation
                {                    
                    workSheet.Cells.LoadFromCollection(DATasks.GlobalautomationTemplateList, true);
                    workSheet.Cells.AutoFitColumns();
                }                          
                


                package.Save();
                return package.GetAsByteArray();
            }
        }

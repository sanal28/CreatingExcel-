private string CreateCSVFile(DataVisualizationModel model)
        {
            string Status = string.Empty;
            string Status1 = string.Empty;
            string FilePath = string.Empty;
            string OfficeName = model.OfficeName;
            string Entity = model.EntityName;
            string PeriodEndDate = model.PeriodEndDate;
            string RecordID = model.RecordID.ToString();
            FilePath = ConfigurationManager.AppSettings["fff"].ToString().Replace("{Office}", OfficeName.Replace(" ", "")).Replace("{Entity-PeriodEndDate}", Entity + '-' + PeriodEndDate);
            var CSVList = GetCSVList(model.ProcessCode);
            var CubeList = GetCubeList(model.ProjectCode);
            StringBuilder sb = new StringBuilder();
            StringBuilder sb1 = new StringBuilder();
            List<string> CsvRow = new List<string>();
            List<string> CubeRow = new List<string>();
            //Add headers

            sb.AppendLine("Tag_Name|Display_Name|Account_Number");

            foreach (DataRow row in CSVList.Rows)
            {
                var tagName = row["CSVList"].ToString().Split('|')[0];
                var DisplayName = row["CSVList"].ToString().Split('|')[1];
                var arrayStr = row["CSVList"].ToString().Split('|')[2];
                var array = arrayStr.Split(';');
                for (var i = 0; i < array.Length; i++)
                {
                    var line = tagName + '|' + DisplayName + '|' + array[i];
                    sb.AppendLine(line.ToString());
                }

            }
            string fileName = ConfigurationManager.AppSettings[""].ToString();
            string FileName = fileName + "_" + RecordID + ".csv";

            string SaveLocation = FilePath + FileName;
            if (File.Exists(SaveLocation))
            {
                Status = "File with Name " + FileName + " already exists.Please change the file name.";
            }
            if (Status == string.Empty)
                File.WriteAllText(SaveLocation, sb.ToString());


            // Cube
            var count = CubeList.Rows.Count;

            if (count != 0)
            {
                sb1.AppendLine("SystemOrManualEntries|SystemJournalEntryIdentificationField|SystemEntryValue");
                foreach (DataRow row in CubeList.Rows)
                {
                    var Entry = row["CSVList"].ToString().Split('|')[0];
                    var Field = row["CSVList"].ToString().Split('|')[1];
                    var arrayStr = row["CSVList"].ToString().Split('|')[2];
                    var array = arrayStr.Split(';');
                    for (var i = 0; i < array.Length; i++)
                    {
                        var line = Entry + '|' + Field + '|' + array[i];
                        sb1.AppendLine(line.ToString());
                    }
                    string CubefileName = ConfigurationManager.AppSettings[""].ToString();
                    string CubeFileName = CubefileName + "_" + RecordID + ".csv";

                    string CubeSaveLocation = FilePath + CubeFileName;
                    if (File.Exists(CubeSaveLocation))
                    {
                        Status1 = "File with Name " + CubeFileName + " already exists.Please change the file name.";
                    }
                    if (Status1 == string.Empty)
                        File.WriteAllText(CubeSaveLocation, sb1.ToString());

                }
            }
            return (Status + Status1);
        }

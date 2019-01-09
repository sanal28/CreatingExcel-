using ETLCenter.CommonLibrary;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using System.Web.UI.WebControls;
using ClosedXML;
using ClosedXML.Excel;
using OfficeOpenXml;

namespace ETLCenter.Controllers
{
    public class ETLReportController : Controller
    {
        GridView aggregateGrid = new GridView();
        DataTable aggregateData = new DataTable();
        [SessionAuthorize]
        public ActionResult Aggregate()
        {
            return View();
        }
        [HttpPost]
        [SessionAuthorize]
        public JsonResult GetAggregateReport(int PackageGroup, int packageName, int EmployeeName, string startDate, string EndDate, int pageNo, int RowCount)
        {
            DateTime? aggregateDate;
            DateTime? aggrgateEndDate;
            if (startDate == "")
            {

                aggregateDate = null;

            }
            else
            {
                aggregateDate = Convert.ToDateTime(startDate);
            }
            if (EndDate == "")
            {

                aggrgateEndDate = null;

            }
            else
            {
                aggrgateEndDate = Convert.ToDateTime(EndDate);
            }
            Session["groupid"] = PackageGroup;
            Session["pkgid"] = packageName;
            Session["empid"] = EmployeeName;
            Session["startdate"] = aggregateDate;
            Session["enddate"] = aggrgateEndDate;
            // DateTime StartDateS;
            // DateTime EndDateS;
            try
            {
                //StartDateS = Convert.ToDateTime(startDate);
                // EndDateS = Convert.ToDateTime(EndDate);
                string AggregateDetails = string.Empty;
                ETLCenterReportService.ETLCenterReportService selectAggregateDetails = new ETLCenterReportService.ETLCenterReportService();
                selectAggregateDetails.Url = Constants.EtlReport;
                AggregateDetails = selectAggregateDetails.GetAggregateReport(pageNo, RowCount, PackageGroup, packageName, EmployeeName, aggregateDate, aggrgateEndDate);
                DataTable aggregatedt = JsonConvert.DeserializeObject<DataTable>(AggregateDetails);
                int Rowcount = Convert.ToInt32(aggregatedt.Rows[0]["TotalRowCount"]);
                Session["AggregateRowCount"] = Convert.ToInt32(Rowcount);
                if (AggregateDetails != string.Empty)
                    return Json(AggregateDetails, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                CommonFunctions commonFun = new CommonFunctions();
                commonFun.ExceptionLog(ControllerContext.HttpContext, ex.Message, ex.TargetSite.Name,
                    Convert.ToString(ControllerContext.RouteData.Values["action"]),
                    Convert.ToString(ControllerContext.RouteData.Values["controller"]));
                commonFun = null;
                return Json(new { flag = false }, JsonRequestBehavior.AllowGet);
            }
            finally
            {
                Dispose();
            }
            return Json(CommonLibrary.Constants.JsonError, JsonRequestBehavior.AllowGet);
        }
        [SessionAuthorize]
        public ActionResult AggregateExport()
        {
            try
            {

                GridView aggregateGrid = new GridView();
                DataTable aggregateData = new DataTable();


                ETLCenterReportService.ETLCenterReportService selectAggregateDetails = new ETLCenterReportService.ETLCenterReportService();
                selectAggregateDetails.Url = Constants.EtlReport;
                string AggregateList = selectAggregateDetails.ExportAggregateReport(Convert.ToInt32(Session["groupid"]), Convert.ToInt32(Session["pkgid"]), Convert.ToInt32(Session["empid"]), Convert.ToDateTime(Session["startdate"]), Convert.ToDateTime(Session["enddate"]), Convert.ToInt32(Session["AggregateRowCount"]));
                if (AggregateList != "")
                {
                    aggregateData = JsonConvert.DeserializeObject<DataTable>(AggregateList);
                    aggregateData.Columns.Remove("TotalRowCount");
                    if (aggregateData.Rows.Count > 0)
                    {
                        aggregateGrid.DataSource = aggregateData;
                        aggregateGrid.DataBind();
                    }
                    if (aggregateGrid.Rows.Count > 0)
                    {
                        //Response.Clear();
                        //Response.AddHeader("content-disposition", "attachment;filename=AggregateReport.xls");
                        //Response.Charset = "";
                        //Response.ContentType = "application/ms-excel";
                        //StringWriter Write = new StringWriter();
                        //HtmlTextWriter HtmlWrite = new HtmlTextWriter(Write);
                        //aggregateGrid.RenderControl(HtmlWrite);
                        //Response.Write(Write.ToString());
                        //Response.End();
                        //Response.Clear();



                        //Response.Clear();
                        //Response.AddHeader("content-disposition", "attachment;filename=AggregateReport.xls");
                        //Response.ContentType = "application/ms-excel";
                        //Response.ContentEncoding = System.Text.Encoding.Unicode;
                        //Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());

                        //System.IO.StringWriter sw = new System.IO.StringWriter();
                        //System.Web.UI.HtmlTextWriter hw = new HtmlTextWriter(sw);

                        //aggregateGrid.RenderControl(hw);

                        //Response.Write(sw.ToString());
                        //Response.End();


                        ExcelPackage excel = new ExcelPackage();
                        var workSheet = excel.Workbook.Worksheets.Add("AggregateReportSheet");

                        //workSheet.Cells["A1"].Style.Numberformat.Format = "dd/MM/yyyy h:mm:ss";
                        workSheet.Cells["A1"].LoadFromDataTable(aggregateData, true);

                        int i = 1;
                        foreach (DataColumn dc in aggregateData.Columns)
                        {

                            if (dc.DataType == typeof(DateTime))
                            {
                                workSheet.Column(i).Style.Numberformat.Format = "MM/dd/yyyy hh:mm:ss";

                            }
                            workSheet.Column(i).AutoFit();
                            i++;
                        }

                        using (var memoryStream = new MemoryStream())
                        {
                            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            Response.AddHeader("content-disposition", "attachment;filename=AggregateReport.xlsx");
                            excel.SaveAs(memoryStream);
                            memoryStream.WriteTo(Response.OutputStream);
                            Response.Flush();
                            Response.End();
                        }


                    }
                    else
                    {
                        return View("Aggregate");

                        //return Json(new { flag = -1 }, JsonRequestBehavior.AllowGet);

                    }
                }
                return View("Aggregate");
            }
            catch (Exception ex)
            {
                CommonFunctions commonFun = new CommonFunctions();
                commonFun.ExceptionLog(ControllerContext.HttpContext, ex.Message, ex.TargetSite.Name,
                    Convert.ToString(ControllerContext.RouteData.Values["action"]),
                    Convert.ToString(ControllerContext.RouteData.Values["controller"]));
                commonFun = null;
                // return null;
                return View("Aggregate");
                // return Json(JsonRequestBehavior.AllowGet);
            }
            finally
            {
                Session["groupid"] = null;
                Session["pkgid"] = null;
                Session["empid"] = null;
                Session["startdate"] = null;
                Session["enddate"] = null;
                Session["AggregateRowCount"] = null;
                Dispose();
            }
        }


        //log report

        [SessionAuthorize]
        public ActionResult Log()
        {
            return View();
        }
        [HttpGet]
        public JsonResult GetPackgestatus()
        {
            ETLCenterReportService.ETLCenterReportService bindpackagestatus = new ETLCenterReportService.ETLCenterReportService();
            bindpackagestatus.Url = Constants.EtlReport;
            try
            {
                string JsonString = bindpackagestatus.BindPackageStatus();
                if (JsonString != string.Empty)
                    return Json(JsonString, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                CommonFunctions commonFun = new CommonFunctions();
                commonFun.ExceptionLog(ControllerContext.HttpContext, ex.Message, ex.TargetSite.Name,
                    Convert.ToString(ControllerContext.RouteData.Values["action"]),
                    Convert.ToString(ControllerContext.RouteData.Values["controller"]));
                commonFun = null;

            }
            finally
            {
                Dispose();

            }
            return Json(CommonLibrary.Constants.JsonError);
        }
        [HttpPost]
        [SessionAuthorize]
        public JsonResult GetLogReport(int PackageGroup, int packageName, int packageStatus, string startDate, string EndDate, int pageNo, int RowCount)
        {
            DateTime? logDate;
            DateTime? logEndDate;
            if (startDate == "")
            {

                logDate = null;

            }
            else
            {
                logDate = Convert.ToDateTime(startDate);
            }
            if (EndDate == "")
            {

                logEndDate = null;

            }
            else
            {
                logEndDate = Convert.ToDateTime(EndDate);
            }
            Session["loggroupid"] = PackageGroup;
            Session["logpkgid"] = packageName;
            Session["logpkgstatus"] = packageStatus;
            Session["logstartdate"] = logDate;
            Session["logenddate"] = logEndDate;

            try
            {

                string logDetails = string.Empty;
                ETLCenterReportService.ETLCenterReportService selectlogDetails = new ETLCenterReportService.ETLCenterReportService();
                selectlogDetails.Url = Constants.EtlReport;
                logDetails = selectlogDetails.GetLogReport(pageNo, RowCount, PackageGroup, packageName, packageStatus, logDate, logEndDate);
                DataTable logData = JsonConvert.DeserializeObject<DataTable>(logDetails);

                int Rowcount = Convert.ToInt32(logData.Rows[0]["TotalRowCount"]);
                Session["logRowCount"] = Convert.ToInt32(Rowcount);

                if (logDetails != string.Empty)
                    return Json(logDetails, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                CommonFunctions commonFun = new CommonFunctions();
                commonFun.ExceptionLog(ControllerContext.HttpContext, ex.Message, ex.TargetSite.Name,
                    Convert.ToString(ControllerContext.RouteData.Values["action"]),
                    Convert.ToString(ControllerContext.RouteData.Values["controller"]));
                commonFun = null;
                return Json(new { flag = false }, JsonRequestBehavior.AllowGet);
            }
            finally
            {
                Dispose();
            }
            return Json(CommonLibrary.Constants.JsonError, JsonRequestBehavior.AllowGet);
        }

        [SessionAuthorize]
        public ActionResult LogExport()
        {
            try
            {
                if (Convert.ToInt32(Session["loggroupid"]) != 0)
                {
                    GridView logGrid = new GridView();
                    DataTable logData = new DataTable();


                    ETLCenterReportService.ETLCenterReportService selectAggregateDetails = new ETLCenterReportService.ETLCenterReportService();
                    selectAggregateDetails.Url = Constants.EtlReport;
                    string LogList = selectAggregateDetails.ExportLogReport(Convert.ToInt32(Session["loggroupid"]), Convert.ToInt32(Session["logpkgid"]), Convert.ToInt32(Session["logpkgstatus"]), Convert.ToDateTime(Session["logstartdate"]), Convert.ToDateTime(Session["logenddate"]), Convert.ToInt32(Session["logRowCount"]));
                    if (LogList != "")
                    {
                        logData = JsonConvert.DeserializeObject<DataTable>(LogList);
                        logData.Columns.Remove("TotalRowCount");
                        //logData.Columns[3].DataType = typeof(DateTime);
                        string colformat = logData.Columns[3].DataType.ToString();
                        if (logData.Rows.Count > 0)
                        {
                            logGrid.DataSource = logData;
                            logGrid.DataBind();
                        }
                        if (logGrid.Rows.Count > 0)
                        {
                            //Response.Clear();
                            //Response.AddHeader("content-disposition", "attachment;filename=LogReport.xls");
                            //Response.Charset = "";
                            //Response.ContentType = "application/ms-excel";
                            //StringWriter Write = new StringWriter();
                            //HtmlTextWriter HtmlWrite = new HtmlTextWriter(Write);
                            //logGrid.RenderControl(HtmlWrite);
                            //Response.Write(Write.ToString());
                            //Response.End();
                            //Response.Clear();




                            //Response.Clear();

                            //Response.AddHeader("content-disposition", "attachment;filename=LogReport.xls");
                            //Response.ContentType = "application/ms-excel";
                            //Response.ContentEncoding = System.Text.Encoding.Unicode;
                            //Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());

                            //System.IO.StringWriter sw = new System.IO.StringWriter();
                            //System.Web.UI.HtmlTextWriter hw = new HtmlTextWriter(sw);

                            //logGrid.RenderControl(hw);

                            //Response.Write(sw.ToString());
                            //Response.End();




                            ExcelPackage excel = new ExcelPackage();
                            var workSheet = excel.Workbook.Worksheets.Add("LogReportSheet");

                            //workSheet.Cells["A1"].Style.Numberformat.Format = "dd/MM/yyyy h:mm:ss";
                            workSheet.Cells["A1"].LoadFromDataTable(logData, true);

                            int i = 1;
                            foreach (DataColumn dc in logData.Columns)
                            {

                                if (dc.DataType == typeof(DateTime))
                                {
                                    workSheet.Column(i).Style.Numberformat.Format = "MM/dd/yyyy hh:mm:ss";

                                }
                                workSheet.Column(i).AutoFit();
                                i++;
                            }

                            using (var memoryStream = new MemoryStream())
                            {
                                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                                Response.AddHeader("content-disposition", "attachment;filename=LogReport.xlsx");
                                excel.SaveAs(memoryStream);
                                memoryStream.WriteTo(Response.OutputStream);
                                Response.Flush();
                                Response.End();
                            }



                            //FileInfo excelFile1 = new FileInfo(@"F:\test2.xlsx");
                            //using (ExcelPackage pck = new ExcelPackage(excelFile1))
                            //{
                            //    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("LogReport1");
                            //    ws.Cells["A1"].LoadFromDataTable(logData, true);
                            //    pck.Save();
                            //}











                            //using (ExcelPackage excel = new ExcelPackage())
                            //{
                            //    //excel.
                            //    excel.Workbook.Worksheets.Add("LogReport");
                            //   // excel.Workbook.Worksheets.Add("Worksheet2");
                            //   // excel.Workbook.Worksheets.Add("Worksheet3");

                            //    var headerRow = new List<string[]>()
                            //      {
                            //        new string[] { "ID", "First Name", "Last Name", "DOB" }
                            //      };

                            //    // Determine the header range (e.g. A1:D1)
                            //    string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

                            //    // Target a worksheet
                            //    var worksheet = excel.Workbook.Worksheets["LogReport"];


                            //    worksheet.Cells[headerRange].LoadFromArrays(headerRow);

                            //    FileInfo excelFile = new FileInfo(@"F:\test.xlsx");

                            //    excel.SaveAs(excelFile);
                            //}









                            //MemoryStream MyMemoryStream = new MemoryStream();
                            //using (XLWorkbook wb = new XLWorkbook())
                            //{
                            // ///   wb.Worksheets.
                            //wb.Worksheets.Add();
                            //    wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            //    wb.Style.Font.Bold = true;

                            //    Response.Clear();
                            //    Response.Buffer = true;
                            //    Response.Charset = "";
                            //    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            //    Response.AddHeader("content-disposition", "attachment;filename= LogReport.xlsx");
                            //    wb.SaveAs(MyMemoryStream);
                            //    MyMemoryStream.WriteTo(Response.OutputStream);
                            //    Response.Flush();
                            //    Response.End();
                            //    using ( MyMemoryStream = new MemoryStream())
                            //    {
                            //        wb.SaveAs(MyMemoryStream);
                            //        MyMemoryStream.WriteTo(Response.OutputStream);
                            //        Response.Flush();
                            //        Response.End();
                            //    }
                            //}

                        }
                        else
                        {
                            return View("Log");

                            //return Json(new { flag = -1 }, JsonRequestBehavior.AllowGet);

                        }
                    }
                    return View("Log");
                }
                else
                {
                    return View("Log");
                }
            }
            catch (Exception ex)
            {
                CommonFunctions commonFun = new CommonFunctions();
                commonFun.ExceptionLog(ControllerContext.HttpContext, ex.Message, ex.TargetSite.Name,
                    Convert.ToString(ControllerContext.RouteData.Values["action"]),
                    Convert.ToString(ControllerContext.RouteData.Values["controller"]));
                commonFun = null;
                // return null;
                return View("Log");

            }
            finally
            {
                Session["loggroupid"] = null;
                Session["logpkgid"] = null;
                Session["logpkgstatus"] = null;
                Session["logstartdate"] = null;
                Session["logenddate"] = null;
                Session["logRowCount"] = null;
                Dispose();

            }
        }

        [SessionAuthorize]
        public ActionResult Error()
        {
            return View();
        }
        [HttpPost]
        [SessionAuthorize]
        public JsonResult GetErrorReport(int PackageGroup, int packageName, string startDate, string EndDate, int pageNo, int RowCount)
        {

            DateTime? logDate;
            DateTime? logEndDate;
            if (startDate == "")
            {

                logDate = null;

            }
            else
            {
                logDate = Convert.ToDateTime(startDate);
            }
            if (EndDate == "")
            {

                logEndDate = null;

            }
            else
            {
                logEndDate = Convert.ToDateTime(EndDate);
            }
            Session["errorgroupid"] = PackageGroup;
            Session["errorogpkgid"] = packageName;
            Session["errorlogstartdate"] = logDate;
            Session["errorlogenddate"] = logEndDate;

            try
            {

                string errorDetails = string.Empty;
                ETLCenterReportService.ETLCenterReportService selecterrorDetails = new ETLCenterReportService.ETLCenterReportService();
                selecterrorDetails.Url = Constants.EtlReport;
                errorDetails = selecterrorDetails.GetErrorReport(pageNo, RowCount, PackageGroup, packageName, logDate, logEndDate);
                DataTable errorData = JsonConvert.DeserializeObject<DataTable>(errorDetails);


                int Rowcount = Convert.ToInt32(errorData.Rows[0]["TotalRowCount"]);
                Session["errorRowCount"] = Convert.ToInt32(Rowcount);
                if (errorDetails != string.Empty)
                    return Json(errorDetails, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                CommonFunctions commonFun = new CommonFunctions();
                commonFun.ExceptionLog(ControllerContext.HttpContext, ex.Message, ex.TargetSite.Name,
                    Convert.ToString(ControllerContext.RouteData.Values["action"]),
                    Convert.ToString(ControllerContext.RouteData.Values["controller"]));
                commonFun = null;
                return Json(new { flag = false }, JsonRequestBehavior.AllowGet);
            }
            finally
            {
                Dispose();
            }
            return Json(CommonLibrary.Constants.JsonError, JsonRequestBehavior.AllowGet);
        }
        [SessionAuthorize]
        public ActionResult ErrorExport()
        {
            try
            {
                if (Convert.ToInt32(Session["errorgroupid"]) != 0)
                {


                    GridView errorGrid = new GridView();
                    DataTable errorData = new DataTable();


                    ETLCenterReportService.ETLCenterReportService selectErrorDetails = new ETLCenterReportService.ETLCenterReportService();
                    selectErrorDetails.Url = Constants.EtlReport;
                    string ErrorList = selectErrorDetails.ExportErrorReport(Convert.ToInt32(Session["errorgroupid"]), Convert.ToInt32(Session["errorogpkgid"]), Convert.ToDateTime(Session["errorlogstartdate"]), Convert.ToDateTime(Session["errorlogenddate"]), Convert.ToInt32(Session["errorRowCount"]));
                    if (ErrorList != "")
                    {
                        errorData = JsonConvert.DeserializeObject<DataTable>(ErrorList);
                        errorData.Columns.Remove("TotalRowCount");
                        if (errorData.Rows.Count > 0)
                        {
                            errorGrid.DataSource = errorData;
                            errorGrid.DataBind();
                        }
                        if (errorGrid.Rows.Count > 0)
                        {
                            //Response.Clear();
                            //Response.AddHeader("content-disposition", "attachment;filename=ErrorReport.xls");
                            //Response.Charset = "";
                            //Response.ContentType = "application/ms-excel";
                            //StringWriter Write = new StringWriter();
                            //HtmlTextWriter HtmlWrite = new HtmlTextWriter(Write);
                            //errorGrid.RenderControl(HtmlWrite);
                            //Response.Write(Write.ToString());
                            //Response.End();

                            ExcelPackage excel = new ExcelPackage();
                            var workSheet = excel.Workbook.Worksheets.Add("ErrorReportSheet");

                            //workSheet.Cells["A1"].Style.Numberformat.Format = "dd/MM/yyyy h:mm:ss";
                            workSheet.Cells["A1"].LoadFromDataTable(errorData, true);

                            int i = 1;
                            foreach (DataColumn dc in errorData.Columns)
                            {

                                if (dc.DataType == typeof(DateTime))
                                {
                                    workSheet.Column(i).Style.Numberformat.Format = "MM/dd/yyyy hh:mm:ss";

                                }
                                workSheet.Column(i).AutoFit();
                                i++;
                            }

                            using (var memoryStream = new MemoryStream())
                            {
                                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                                Response.AddHeader("content-disposition", "attachment;filename=ErrorReport.xlsx");
                                excel.SaveAs(memoryStream);
                                memoryStream.WriteTo(Response.OutputStream);
                                Response.Flush();
                                Response.End();
                            }


                            //Response.Clear();
                            //Response.AddHeader("content-disposition", "attachment;filename=ErrorReport.xls");
                            //Response.ContentType = "application/ms-excel";
                            //Response.ContentEncoding = System.Text.Encoding.Unicode;
                            //Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());

                            //System.IO.StringWriter sw = new System.IO.StringWriter();
                            //System.Web.UI.HtmlTextWriter hw = new HtmlTextWriter(sw);

                            //errorGrid.RenderControl(hw);

                            //Response.Write(sw.ToString());
                            //Response.End();


                        }
                        else
                        {
                            return View("Error");

                            //return Json(new { flag = -1 }, JsonRequestBehavior.AllowGet);

                        }
                    }
                    return View("Error");
                }
                else
                {
                    return View("Error");
                }
            }
            catch (Exception ex)
            {
                CommonFunctions commonFun = new CommonFunctions();
                commonFun.ExceptionLog(ControllerContext.HttpContext, ex.Message, ex.TargetSite.Name,
                    Convert.ToString(ControllerContext.RouteData.Values["action"]),
                    Convert.ToString(ControllerContext.RouteData.Values["controller"]));
                commonFun = null;
                // return null;
                return View("Error");
                // return Json(JsonRequestBehavior.AllowGet);
            }
            finally
            {
                Session["errorgroupid"] = null;
                Session["errorogpkgid"] = null;
                Session["errorlogstartdate"] = null;
                Session["errorlogenddate"] = null;
                Session["errorRowCount"] = null;
                Dispose();
            }
        }

        [SessionAuthorize]
        public ActionResult PackageStatistics()
        {
            return View();
        }
        [HttpPost]
        [SessionAuthorize]
        public JsonResult GetpackageReport(int packageName, string startDate, string EndDate, int pageNo, int RowCount)
        {

            DateTime? logDate;
            DateTime? logEndDate;
            if (startDate == "")
            {

                logDate = null;

            }
            else
            {
                logDate = Convert.ToDateTime(startDate);
            }
            if (EndDate == "")
            {

                logEndDate = null;

            }
            else
            {
                logEndDate = Convert.ToDateTime(EndDate);
            }

            //Session["errorogpkgid"] = packageName;
            //Session["errorlogstartdate"] = logDate;
            //Session["errorlogenddate"] = logEndDate;

            try
            {

                string jsonString = string.Empty;
                ETLCenterReportService.ETLCenterReportService selecterrorDetails = new ETLCenterReportService.ETLCenterReportService();
                selecterrorDetails.Url = Constants.EtlReport;

                jsonString = selecterrorDetails.PackagestatisticsGraph(packageName, logDate, logEndDate, pageNo, RowCount); //change value of offset as per pagination case
                if (jsonString != string.Empty)
                    return Json(jsonString, JsonRequestBehavior.AllowGet);

                DataTable logData = JsonConvert.DeserializeObject<DataTable>(jsonString);
                int rowcount = logData.Rows.Count;

            }
            catch (Exception ex)
            {
                CommonFunctions commonFun = new CommonFunctions();
                commonFun.ExceptionLog(ControllerContext.HttpContext, ex.Message, ex.TargetSite.Name,
                    Convert.ToString(ControllerContext.RouteData.Values["action"]),
                    Convert.ToString(ControllerContext.RouteData.Values["controller"]));
                commonFun = null;
                return Json(new { flag = false }, JsonRequestBehavior.AllowGet);
            }
            finally
            {
                Dispose();
            }
            return Json(CommonLibrary.Constants.JsonError, JsonRequestBehavior.AllowGet);
        }

    }
}

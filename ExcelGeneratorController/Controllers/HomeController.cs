using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Services.Description;
using System.Data.SqlClient;
using Syncfusion.XlsIO;
using System.Runtime.InteropServices.ComTypes;

namespace ExcelGeneratorController.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public JsonResult UploadExcel()
        {
            string Result = string.Empty;
            HttpFileCollectionBase httpFile = Request.Files;
            HttpPostedFileBase file = Request.Files[0];
            IExcelDataReader reader = null;
            string path = Server.MapPath("~/Uploaded/");
            string FileName = string.Empty;
            DirectoryInfo di = new DirectoryInfo(path);
            try
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                FileName = file.FileName;
                string filepath = Path.Combine(path, FileName);

                file.SaveAs(filepath);

                using (FileStream stream = System.IO.File.Open(filepath, FileMode.Open, FileAccess.Read))
                {
                    if (file.FileName.EndsWith(".xls"))
                    {
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    }
                    else if (file.FileName.EndsWith(".xlsx"))
                    {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    }
                    else
                    {
                        Result = "This file format is not supported";
                    }

                    DataSet excelRecords = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });
                    reader.Close();

                    DataTable DT = excelRecords.Tables[0];
                    SqlConnection con = new SqlConnection("Data Source = HRMPC545\\SQLEXPRESS; Initial Catalog = Practice; Integrated Security = True");
                    using(SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                    {
                        con.Open();
                        bulkCopy.DestinationTableName = "Users";
                        bulkCopy.WriteToServer(DT);
                    }
                }
            }
            catch(Exception ex)
            {
                Result = ex.Message;
            }
            finally{

            }
            
            return Json(Result, JsonRequestBehavior.AllowGet);
        }

        public JsonResult ExportExcel(Stream stream, DataTable DT,string filepath)
        {
            string Result = string.Empty;
            using (ExcelEngine excel = new ExcelEngine())
            {
                IApplication application = excel.Excel;
                application.DefaultVersion = ExcelVersion.Excel2013;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];
                worksheet.ImportDataTable(DT, true, 1, 1, true);
                worksheet.UsedRange.AutofitColumns();
                workbook.SaveAs(stream);
            }

            byte[] bytes = System.IO.File.ReadAllBytes(filepath);
            System.IO.File.Delete(filepath);
            return Json(Convert.ToBase64String(bytes),JsonRequestBehavior.AllowGet);
        }
    }
}
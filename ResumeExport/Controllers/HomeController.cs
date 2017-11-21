using ResumeExport.Service;
using System;
using System.Web.Mvc;

namespace ResumeExport.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {   
            return View();
        }

        public ActionResult SpireDoc_ExportByHtml()
        {
            bool result;
            string msg;
            byte[] objFile = new SpireDocExportService().ExportResumeByHtml(out result, out msg);

            if (result)
            {
                ////Word (doc)
                //return File(objFile, "application/msword", "MyReseme.doc");
                ////PDF
                //return File(objFile, "application/pdf", "MyReseme.pdf");
                //Word (docx)
                return File(objFile, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "MyResume.docx");
            }
            else
            {
                return RedirectToAction("Index");
            }
        }
        
        public ActionResult SpireDoc_ExportResumeByDocx()
        {
            bool result;
            string msg;
            byte[] objFile = new SpireDocExportService().ExportResumeByDocx_ReplaceText(out result, out msg);

            if (result)
            {
                ////Word (doc)
                //return File(objFile, "application/msword", "MyReseme.doc");
                ////PDF
                //return File(objFile, "application/pdf", "MyReseme.pdf");
                ////Word (docx)
                return File(objFile, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "MyResume.docx");
            }
            else
            {
                return RedirectToAction("Index");
            }
        }
        
        public ActionResult SpireDoc_ExportResumeByDocxBookmark()
        {
            bool result;
            string msg;
            byte[] objFile = new SpireDocExportService().ExportResumeByDocx_Bookmark(out result, out msg);

            if (result)
            {
                ////Word (doc)
                //return File(objFile, "application/msword", "MyReseme.doc");
                ////PDF
                //return File(objFile, "application/pdf", "MyReseme.pdf");
                ////Word (docx)
                return File(objFile, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "MyResume.docx");
            }
            else
            {
                return RedirectToAction("Index");
            }
        }
        
        public ActionResult SpireDoc_ExportResumePdfByDocx()
        {
            bool result;
            string msg;
            byte[] objFile = new SpireDocExportService().ExportResume_Word2PDF(out result, out msg);

            if (result)
            {
                ////Word (doc)
                //return File(objFile, "application/msword", "MyReseme.doc");
                ////PDF
                //return File(objFile, "application/pdf", "MyReseme.pdf");
                ////Word (docx)
                return File(objFile, System.Net.Mime.MediaTypeNames.Application.Pdf, "MyResume.pdf");
            }
            else
            {
                return RedirectToAction("Index");
            }
        }

        public ActionResult OpenXML_ExportByHtml()
        {
            bool result;
            string msg;
            byte[] objFile = new OpenXmlExportService().ExportByHtml(out result, out msg);

            if (result)
            {
                ////Word (doc)
                //return File(objFile, "application/msword", "MyReseme.doc");
                ////PDF
                //return File(objFile, "application/pdf", "MyReseme.pdf");
                //Word (docx)
                return File(objFile, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "MyResume.docx");
            }
            else
            {
                return RedirectToAction("Index");
            }
        }

        public ActionResult OpenXML_ExportByDocx()
        {
            bool result;
            string msg;
            byte[] objFile = new OpenXmlExportService().ExportByDocx(out result, out msg);

            if (result)
            {
                ////Word (doc)
                //return File(objFile, "application/msword", "MyReseme.doc");
                ////PDF
                //return File(objFile, "application/pdf", "MyReseme.pdf");
                //Word (docx)
                return File(objFile, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "MyResume.docx");
            }
            else
            {
                return RedirectToAction("Index");
            }
        }

        public ActionResult DocX_ExportImgDocx()
        {
            byte[] objFile = DocXService.AddImgDoc("https://github.com/felixhuang72/Sample_ExportToDocx/raw/master/MarkdownAttach/iis_officedom_01.png");
            if(objFile != null && objFile.Length > 0)
            {
                return File(objFile, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"docx_test_{DateTime.Now.ToString("yyyyMMddHHmmss")}.docx");
            }
            else
            {
                return RedirectToAction("Index");
            }
        }
        



        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}
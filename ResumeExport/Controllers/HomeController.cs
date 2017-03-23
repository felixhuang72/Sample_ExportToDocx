using ResumeExport.Service;
using System.Web.Mvc;

namespace ResumeExport.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ExportByHtml()
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
                return File(objFile, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "MyReseme.docx");
            }
            else
            {
                return RedirectToAction("Index");
            }
        }


        public ActionResult ExportResumeByDocx()
        {
            bool result;
            string msg;
            byte[] objFile = new SpireDocExportService().ExportResumeByDocx(out result, out msg);

            if (result)
            {
                ////Word (doc)
                //return File(objFile, "application/msword", "MyReseme.doc");
                ////PDF
                //return File(objFile, "application/pdf", "MyReseme.pdf");
                ////Word (docx)
                return File(objFile, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "MyReseme.docx");
            }
            else
            {
                return RedirectToAction("Index");
            }
        }


        public ActionResult Docx_Export()
        {
            bool result;
            string msg;
            byte[] objFile = new DocXExportService().Export(out result, out msg);

            if (result)
            {
                ////Word (doc)
                //return File(objFile, "application/msword", "MyReseme.doc");
                ////PDF
                //return File(objFile, "application/pdf", "MyReseme.pdf");
                //Word (docx)
                return File(objFile, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "MyReseme.docx");
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
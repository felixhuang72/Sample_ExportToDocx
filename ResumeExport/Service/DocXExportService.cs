using Novacode;
using System;
using System.IO;

namespace ResumeExport.Service
{
    //檔案匯出服務 (DocX)
    public class DocXExportService
    {
        public byte[] Export(out bool result, out string msg)
        {
            result = true;
            msg = "";
            MemoryStream ms = new MemoryStream();

            try
            {
                using (DocX doc = DocX.Create("Example.docx"))
                {
                    Novacode.Paragraph p = doc.InsertParagraph("Hello test");
                    doc.SaveAs(ms);
                }
            }
            catch (Exception ex)
            {
                result = false;
                msg = ex.Message;
            }


            if (result)
            {
                return ms.ToArray();
            }
            else
            {
                return null;
            }
        }
    }
}
using Aspose.Words;
using Microsoft.Office.Interop.Word;
using Spire.Doc;
using System;
using System.IO;
using System.Web;

namespace ResumeExport.Service
{
    /// <summary>
    /// 將 Word 檔案轉存成 PDF
    /// </summary>
    public class ConvertPDFService
    {
        /****** 程式處理邏輯 *****
        * 
        * 1. 取得 docx 串流資訊，將其轉換為實體暫存檔案
        * 2. 透過 Microsoft.Office.Interop.Word 套件將方才產生的 docx 暫存檔轉換為 PDF 實體檔案
        * 3. 將 PDF 實體檔案轉換成串流資訊
        * 4. 回傳前，將產生的暫存檔案 (docx, pdf) 移除
        * 5. 回傳 PDF 串流資訊，完成
        * 
        ************************/


        /// <summary>
        /// 將 Word 串流轉換為 PDF 串流 (使用 Microsoft.Office.Interop.Word 套件)
        /// </summary>
        /// <param name="WordFileStreaming">Word 串流內容</param>
        /// <returns>PDF串流</returns>
        public static byte[] ConvertToPdf_MicrosoftOfficeInteropWord(byte[] WordFileStreaming)
        {
            bool result = true;
            string msg = "";
            MemoryStream ms = new MemoryStream();
            FileStream fs = null;

            //利用 Spire.Doc 將 Word 串流資訊轉換為實體暫存檔案
            Spire.Doc.Document spiredoc = new Spire.Doc.Document();
            Stream tmpdoc = new MemoryStream(WordFileStreaming);
            spiredoc.LoadFromStream(tmpdoc, FileFormat.Docx);

            string tmpDocDir = HttpContext.Current.Server.MapPath("~/TmpDocs");
            string tmpDocPath = Path.Combine(tmpDocDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx");
            if (!Directory.Exists(tmpDocDir))
            {
                Directory.CreateDirectory(tmpDocDir);
            }
            spiredoc.SaveToFile(tmpDocPath, FileFormat.Docx);

            //轉換成 PDF
            string tmpPdfFilePath = Path.Combine(tmpDocDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf");
            if (File.Exists(tmpDocPath))
            {
                var appWord = new Application();
                if (appWord.Documents != null)
                {
                    var wordDocument = appWord.Documents.Open(tmpDocPath);
                    if (wordDocument != null)
                    {
                        try
                        {
                            //將 Word 檔轉存成 PDF 實體檔案
                            wordDocument.ExportAsFixedFormat(tmpPdfFilePath, WdExportFormat.wdExportFormatPDF);
                            //將轉換後的 PDF 實體檔案串流化
                            fs = new FileStream(tmpPdfFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                            wordDocument.Close();
                            //將 FileStream 轉存給 MemoryStream
                            fs.CopyTo(ms);
                        }
                        catch (Exception ex) { result = false; msg = ex.Message; }
                        finally
                        {
                            fs.Dispose();
                            //刪除產生的暫存 PDF 檔
                            File.Delete(tmpPdfFilePath);
                        }
                    }
                }
                appWord.Quit();

                //刪除產生的暫存 Word 檔
                File.Delete(tmpDocPath);
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


        /// <summary>
        /// 將 Word 串流轉換為 PDF 串流(使用 Aspose.Words 套件)
        /// </summary>
        /// <param name="WordFileStreaming">Word 串流內容</param>
        /// <returns>PDF串流</returns>
        public static byte[] ConvertToPdf_AsposeWords(byte[] WordFileStreaming)
        {
            bool result = true;
            string msg = "";
            MemoryStream ms = new MemoryStream();
            FileStream fs = null;


            //利用 Spire.Doc 將 Word 串流資訊轉換為實體暫存檔案
            Spire.Doc.Document spiredoc = new Spire.Doc.Document();
            Stream tmpdoc = new MemoryStream(WordFileStreaming);
            spiredoc.LoadFromStream(tmpdoc, FileFormat.Docx);

            string tmpDocDir = HttpContext.Current.Server.MapPath("~/TmpDocs");
            string tmpDocPath = Path.Combine(tmpDocDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx");
            if (!Directory.Exists(tmpDocDir))
            {
                Directory.CreateDirectory(tmpDocDir);
            }
            spiredoc.SaveToFile(tmpDocPath, FileFormat.Docx);


            //轉換成 PDF
            string tmpPdfFilePath = Path.Combine(tmpDocDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf");
            if (File.Exists(tmpDocPath))
            {
                try
                {
                    //轉換過程
                    Aspose.Words.Document srcDoc = new Aspose.Words.Document(tmpDocPath);
                    // Save the document in PDF format.
                    srcDoc.Save(tmpPdfFilePath, SaveFormat.Pdf);

                    //將轉換後的 PDF 實體檔案串流化
                    fs = new FileStream(tmpPdfFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    //將 FileStream 轉存給 MemoryStream
                    fs.CopyTo(ms);

                }
                catch (Exception ex)
                {
                    result = false;
                    msg = ex.Message;
                }
                finally
                {
                    fs.Dispose();
                    //刪除產生的暫存 PDF 檔
                    File.Delete(tmpPdfFilePath);
                }

                //刪除產生的暫存 Word 檔
                File.Delete(tmpDocPath);
            }


            // 回傳串流資訊
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
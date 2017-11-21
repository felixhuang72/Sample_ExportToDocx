using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using Xceed.Words.NET;

namespace ResumeExport.Service
{
    /// <summary>
    /// DocX 相關操作測試
    /// </summary>
    public class DocXService
    {
        /// <summary>
        /// 透過 WebClient 方式取得遠端圖片，插入並匯出 docx 檔案
        /// </summary>
        /// <param name="imgSourceUrl">遠端圖片位置</param>
        /// <returns></returns>
        public static byte[] AddImgDoc(string imgSourceUrl)
        {
            byte[] img = null;
            MemoryStream ms = new MemoryStream();

            //自遠端取得圖片
            try
            {
                WebClient ws = new WebClient();
                img = ws.DownloadData(imgSourceUrl);
            }
            catch (Exception ex) { }

            //若有 Request 到圖片，建立 docx 文件內容
            if (img != null && img.Length > 0)
            {
                using (DocX document = DocX.Create($"DocXTest_{DateTime.Now.ToString("yyyyMMddHHmmss")}.docx"))
                {
                    //建立圖片物件
                    Image image = document.AddImage(new MemoryStream(img));
                    Picture picture = image.CreatePicture();
                    //picture.Rotation = 10;
                    //picture.SetPictureShape(BasicShapes.cube);

                    //建立文件內容
                    Paragraph title = document.InsertParagraph().Append("This is a test for a picture").FontSize(20);
                    title.Alignment = Alignment.center;

                    // Insert a new Paragraph into the document.
                    Paragraph p1 = document.InsertParagraph();

                    // Append content to the Paragraph
                    p1.AppendLine("Just below there should be a picture ").Append("picture").Bold().Append(" inserted in a non-conventional way.");
                    p1.AppendLine();
                    p1.AppendLine("Check out this picture ").AppendPicture(picture).Append(" its funky don't you think?");
                    p1.AppendLine();

                    document.SaveAs(ms);
                }
                return ms.ToArray();
            }
            else
            {
                return null;
            }
        }
    }
}
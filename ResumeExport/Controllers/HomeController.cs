using ResumeExport.Models;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace ResumeExport.Controllers
{
    #region Controller

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
            byte[] objFile = new FileService().ExportResumeByHtml(out result, out msg);

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
            byte[] objFile = new FileService().ExportResumeByDocx(out result, out msg);

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

    #endregion


    #region Service

    //檔案匯出服務
    public class FileService
    {
        /// <summary>
        /// 透過 HTML Tag 匯出 Word 文件
        /// </summary>
        /// <param name="result">回傳: 執行結果</param>
        /// <param name="msg">回傳: 訊息</param>
        /// <returns>串流資訊</returns>
        public byte[] ExportResumeByHtml(out bool result, out string msg)
        {
            result = true;
            msg = "";
            MemoryStream ms = new MemoryStream();

            try
            {
                Spire.Doc.Document document = new Spire.Doc.Document();

                #region 文件內容

                //建立/取得要匯出的內容
                Resume model = new Resume();
                StringBuilder html = new StringBuilder();
                html.Append("Name: " + model.Name + "<br />");
                html.Append("Gender: " + model.Gender + "<br />");
                html.Append("Email: " + model.Email + "<br />");
                html.Append("Address: " + model.Address + "<br />");
                html.Append("Phone: " + model.Phone + "<br />");
                html.Append("Mobile: " + model.Mobile + "<br />");
                html.Append("Description1:<br />" + HttpUtility.HtmlDecode(model.Description1) + "<br /></p>");
                html.Append("Description2:<br />" + HttpUtility.HtmlDecode(model.Description2) + "<br /></p>");

                if (model.JobHistory.Count > 0)
                {
                    int i = 1;
                    model.JobHistory = model.JobHistory.OrderBy(x => x.StartDT).ToList();
                    html.Append("<p>簡歷</p>");
                    html.Append("<table><tr><th>項目</th><th>任職</th><th>職稱</th><th>開始時間</th><th>結束時間</th></tr>");
                    foreach (var h in model.JobHistory)
                    {
                        html.Append("<tr>");
                        html.Append("<td>" + i.ToString() + "</td>");
                        html.Append("<td>" + h.CompanyName + "</td>");
                        html.Append("<td>" + h.JobTitle + "</td>");
                        html.Append("<td>" + (h.StartDT.HasValue ? h.StartDT.Value.ToShortDateString() : "") + "</td>");
                        html.Append("<td>" + (h.EndDT.HasValue ? h.EndDT.Value.ToShortDateString() : "") + "</td>");
                        html.Append("</tr>");
                        i++;
                    }
                    html.Append("</table>");
                }

                #endregion

                //將 HTML 載入至 Document
                document.LoadHTML(new StringReader(html.ToString()), XHTMLValidationType.None);

                #region 設定樣式

                //一般段落文字
                ParagraphStyle style = new ParagraphStyle(document)
                {
                    Name = "BasicStyle"
                };
                //style.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Justify;
                style.CharacterFormat.FontName = "標楷體";
                style.CharacterFormat.FontSize = 12;
                document.Styles.Add(style);

                #endregion

                #region 套用樣式

                for (int s = 0; s < document.Sections.Count; s++)
                {
                    Section section = document.Sections[s];
                    //套用文章段落樣式
                    for (int p = 0; p < section.Paragraphs.Count; p++)
                    {
                        Paragraph pgh = section.Paragraphs[p];
                        pgh.ApplyStyle("BasicStyle");
                        pgh.Format.BeforeSpacing = 10;
                    }

                    //套用表格樣式
                    for (int t = 0; t < document.Sections[s].Tables.Count; t++)
                    {
                        Table table = (Table)document.Sections[s].Tables[t];
                        table.PreferredWidth = new PreferredWidth(WidthType.Percentage, 100);
                        table.TableFormat.IsAutoResized = true;

                        //set table border
                        table.TableFormat.Borders.Right.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                        table.TableFormat.Borders.Left.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                        table.TableFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                        table.TableFormat.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                        table.TableFormat.Borders.Horizontal.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                        table.TableFormat.Borders.Vertical.BorderType = Spire.Doc.Documents.BorderStyle.Thick;

                        for (int tr = 0; tr < table.Rows.Count; tr++)
                        {
                            for (int td = 0; td < table.Rows[tr].Cells.Count; td++)
                            {
                                for (int t_ph = 0; t_ph < table.Rows[tr].Cells[td].Paragraphs.Count; t_ph++)
                                {
                                    table.Rows[tr].Cells[td].Paragraphs[t_ph].ApplyStyle("BasicStyle");
                                }
                            }
                        }
                    }
                }

                #endregion

                //匯出
                document.SaveToStream(ms, FileFormat.Docx);

            }
            catch (Exception ex)
            {
                msg = ex.Message;
                result = false;
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
        /// 透過既有的套印檔匯出 Word 文件
        /// </summary>
        /// <param name="result">回傳: 執行結果</param>
        /// <param name="msg">回傳: 訊息</param>
        /// <returns>串流資訊</returns>
        public byte[] ExportResumeByDocx(out bool result, out string msg)
        {
            result = true;
            msg = "";
            MemoryStream ms = new MemoryStream();

            try
            {
                Spire.Doc.Document document = new Spire.Doc.Document();

                //載入套印檔
                //注意: 實際運作時，若同一時間有兩位以上使用者同時進行套印，會產生「無法開啟已開啟檔案」的錯誤
                //建議實作時，一個使用者執行匯出動作時先複製一個套印檔，完成套印後再將複製的檔案刪除，即可避開錯誤
                document.LoadFromFile(HttpContext.Current.Server.MapPath("~/App_Data/MyResumeSample.docx"));

                #region 定義樣式

                //定義樣式 BasicStyle: 一般段落文字
                ParagraphStyle style = new ParagraphStyle(document)
                {
                    Name = "Basic"
                };
                //style.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Justify;
                style.CharacterFormat.FontName = "標楷體";
                style.CharacterFormat.FontSize = 12;
                document.Styles.Add(style);

                #endregion

                //取得要套印的內容
                Resume model = new Resume();

                #region 套印內容

                document.Replace("{$Name$}", string.IsNullOrEmpty(model.Name) ? "" : model.Name, false, true);
                document.Replace("{$Gender$}", string.IsNullOrEmpty(model.Gender) ? "" : model.Gender, false, true);
                document.Replace("{$Email$}", string.IsNullOrEmpty(model.Email) ? "" : model.Email, false, true);
                document.Replace("{$Address$}", string.IsNullOrEmpty(model.Address) ? "" : model.Address, false, true);
                document.Replace("{$Phone$}", string.IsNullOrEmpty(model.Phone) ? "" : model.Phone, false, true);
                document.Replace("{$Mobile$}", string.IsNullOrEmpty(model.Mobile) ? "" : model.Mobile, false, true);

                //包含 HTML 字串需放置在 paragraph 內，
                //因此套印檔中的 {$Description1$} 及 {$Description2$} 需透過「以 paragraph 取代文字」方式替代
                //Replace {$Description1$} with paragraph
                TextSelection selection = document.FindString("{$Description1$}", false, true);
                TextRange range = selection.GetAsOneRange();
                Paragraph paragraph = range.OwnerParagraph;
                paragraph.ApplyStyle("Basic");
                paragraph.Replace("{$Description1$}", "", false, false);
                paragraph.AppendHTML(string.IsNullOrEmpty(model.Description1) ? "" : HttpUtility.HtmlDecode(model.Description1));

                //Replace {$Description2$} with paragraph
                selection = document.FindString("{$Description2$}", false, true);
                range = selection.GetAsOneRange();
                paragraph = range.OwnerParagraph;
                paragraph.ApplyStyle("Basic");
                paragraph.Replace("{$Description2$}", "", false, false);
                paragraph.AppendHTML(string.IsNullOrEmpty(model.Description2) ? "" : HttpUtility.HtmlDecode(model.Description2));

                #endregion

                #region 動態新增表格

                if (model.JobHistory.Count > 0)
                {
                    Section s = document.AddSection();
                    Table table = s.AddTable(true);
                    string[] Header = { "序號", "任職公司", "職稱", "開始時間", "結束時間" };

                    //Add Cells
                    table.ResetCells(model.JobHistory.Count + 1, Header.Length);

                    //Header Row
                    TableRow FRow = table.Rows[0];
                    FRow.IsHeader = true;
                    for (int i = 0; i < Header.Length; i++)
                    {
                        Paragraph p = FRow.Cells[i].AddParagraph();
                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p.Format.HorizontalAlignment = HorizontalAlignment.Center;

                        TextRange TR = p.AppendText(Header[i]);
                        TR.CharacterFormat.Bold = true;
                    }

                    //Data Row
                    model.JobHistory = model.JobHistory.OrderBy(x => x.StartDT).ToList();
                    for (int r = 0; r < model.JobHistory.Count; r++)
                    {
                        TableRow DataRow = table.Rows[r + 1];
                        string[] data = new string[] { (r + 1).ToString(), model.JobHistory[r].CompanyName, model.JobHistory[r].JobTitle, (model.JobHistory[r].StartDT.HasValue ? model.JobHistory[r].StartDT.Value.ToShortDateString() : ""), (model.JobHistory[r].EndDT.HasValue ? model.JobHistory[r].EndDT.Value.ToShortDateString() : "") };

                        //Columns.
                        for (int c = 0; c < data.Length; c++)
                        {
                            //Cell Alignment
                            DataRow.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                            //Fill Data in Rows
                            Paragraph p2 = DataRow.Cells[c].AddParagraph();
                            TextRange TR2 = p2.AppendText(data[c]);

                            //Format Cells
                            p2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        }
                    }

                    //Replace text with Table
                    TextSelection selectionTable = document.FindString("{$JobHistory$}", true, true);
                    TextRange rangeTable = selectionTable.GetAsOneRange();
                    Paragraph paragraphTable = rangeTable.OwnerParagraph;
                    Body body = paragraphTable.OwnerTextBody;
                    int index = body.ChildObjects.IndexOf(paragraphTable);
                    body.ChildObjects.Remove(paragraphTable);
                    body.ChildObjects.Insert(index, table);
                }

                #endregion

                #region 套用樣式

                //套用文章段落樣式
                for (int s = 0; s < document.Sections.Count; s++)
                {
                    Section section = document.Sections[s];
                    //套用文章段落樣式
                    for (int p = 0; p < section.Paragraphs.Count; p++)
                    {
                        Paragraph pgh = section.Paragraphs[p];
                        pgh.ApplyStyle("Basic");
                        pgh.Format.BeforeSpacing = 12;
                    }

                    //套用表格樣式
                    for (int t = 0; t < document.Sections[s].Tables.Count; t++)
                    {
                        Table table = (Table)document.Sections[s].Tables[t];
                        table.PreferredWidth = new PreferredWidth(WidthType.Percentage, 100);
                        table.TableFormat.IsAutoResized = true;

                        //set table border
                        //table.TableFormat.Borders.Right.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                        //table.TableFormat.Borders.Left.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                        //table.TableFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                        //table.TableFormat.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                        //table.TableFormat.Borders.Horizontal.BorderType = Spire.Doc.Documents.BorderStyle.Thick;
                        //table.TableFormat.Borders.Vertical.BorderType = Spire.Doc.Documents.BorderStyle.Thick;

                        for (int tr = 0; tr < table.Rows.Count; tr++)
                        {
                            for (int td = 0; td < table.Rows[tr].Cells.Count; td++)
                            {
                                for (int t_ph = 0; t_ph < table.Rows[tr].Cells[td].Paragraphs.Count; t_ph++)
                                {
                                    table.Rows[tr].Cells[td].Paragraphs[t_ph].ApplyStyle("Basic");
                                }
                            }
                        }
                    }
                }

                #endregion

                //匯出
                document.SaveToStream(ms, FileFormat.Docx);
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

    #endregion
}
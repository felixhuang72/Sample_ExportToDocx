using ResumeExport.Models;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using Microsoft.Office.Interop.Word;
using System.Drawing;

namespace ResumeExport.Service
{
    //檔案匯出服務 (使用 SpireDoc)
    public class SpireDocExportService
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
                    Spire.Doc.Section section = document.Sections[s];
                    //套用文章段落樣式
                    for (int p = 0; p < section.Paragraphs.Count; p++)
                    {
                        Spire.Doc.Documents.Paragraph pgh = section.Paragraphs[p];
                        pgh.ApplyStyle("BasicStyle");
                        pgh.Format.BeforeSpacing = 10;
                    }

                    //套用表格樣式
                    for (int t = 0; t < document.Sections[s].Tables.Count; t++)
                    {
                        Spire.Doc.Table table = (Spire.Doc.Table)document.Sections[s].Tables[t];
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
        /// 透過既有的套印檔匯出 Word 文件 (以「取代文字」方式套印)
        /// </summary>
        /// <param name="result">回傳: 執行結果</param>
        /// <param name="msg">回傳: 訊息</param>
        /// <returns>串流資訊</returns>
        public byte[] ExportResumeByDocx_ReplaceText(out bool result, out string msg)
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
                Spire.Doc.Documents.Paragraph paragraph = range.OwnerParagraph;
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

                //Replace {$Img$} with Image                
                DocPicture pic = new DocPicture(document);
                pic.LoadImage(Image.FromFile(HttpContext.Current.Server.MapPath("~/App_Data/Penguins.jpg")));

                selection = document.FindString("{$Img$}", false, true);
                range = selection.GetAsOneRange();                
                range.OwnerParagraph.ChildObjects.Insert(0, pic);
                range.OwnerParagraph.ChildObjects.Remove(range);
                
                #endregion

                #region 動態新增表格

                if (model.JobHistory.Count > 0)
                {
                    Spire.Doc.Section s = document.AddSection();
                    Spire.Doc.Table table = s.AddTable(true);
                    string[] Header = { "序號", "任職公司", "職稱", "開始時間", "結束時間" };

                    //Add Cells
                    table.ResetCells(model.JobHistory.Count + 1, Header.Length);

                    //Header Row
                    TableRow FRow = table.Rows[0];
                    FRow.IsHeader = true;
                    for (int i = 0; i < Header.Length; i++)
                    {
                        Spire.Doc.Documents.Paragraph p = FRow.Cells[i].AddParagraph();
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
                            Spire.Doc.Documents.Paragraph p2 = DataRow.Cells[c].AddParagraph();
                            TextRange TR2 = p2.AppendText(data[c]);

                            //Format Cells
                            p2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        }
                    }

                    //Replace text with Table
                    TextSelection selectionTable = document.FindString("{$JobHistory$}", true, true);
                    TextRange rangeTable = selectionTable.GetAsOneRange();
                    Spire.Doc.Documents.Paragraph paragraphTable = rangeTable.OwnerParagraph;
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
                    Spire.Doc.Section section = document.Sections[s];
                    //套用文章段落樣式
                    for (int p = 0; p < section.Paragraphs.Count; p++)
                    {
                        Spire.Doc.Documents.Paragraph pgh = section.Paragraphs[p];
                        pgh.ApplyStyle("Basic");
                        pgh.Format.BeforeSpacing = 12;
                    }

                    //套用表格樣式
                    for (int t = 0; t < document.Sections[s].Tables.Count; t++)
                    {
                        Spire.Doc.Table table = (Spire.Doc.Table)document.Sections[s].Tables[t];
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


        /// <summary>
        /// 透過既有的套印檔匯出 Word 文件 (以「編輯書籤內容」方式套印)
        /// </summary>
        /// <param name="result">回傳: 執行結果</param>
        /// <param name="msg">回傳: 訊息</param>
        /// <returns>串流資訊</returns>
        public byte[] ExportResumeByDocx_Bookmark(out bool result, out string msg)
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
                document.LoadFromFile(HttpContext.Current.Server.MapPath("~/App_Data/MyResumeSample_Bookmark.docx"));

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

                BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);

                bookmarkNavigator.MoveToBookmark("NAME");
                bookmarkNavigator.ReplaceBookmarkContent(string.IsNullOrEmpty(model.Name) ? "" : model.Name, false);
                bookmarkNavigator.MoveToBookmark("GENDER");
                bookmarkNavigator.ReplaceBookmarkContent(string.IsNullOrEmpty(model.Gender) ? "" : model.Gender, false);
                bookmarkNavigator.MoveToBookmark("EMAIL");
                bookmarkNavigator.ReplaceBookmarkContent(string.IsNullOrEmpty(model.Email) ? "" : model.Email, false);
                bookmarkNavigator.MoveToBookmark("ADDRESS");
                bookmarkNavigator.ReplaceBookmarkContent(string.IsNullOrEmpty(model.Address) ? "" : model.Address, false);
                bookmarkNavigator.MoveToBookmark("PHONE");
                bookmarkNavigator.ReplaceBookmarkContent(string.IsNullOrEmpty(model.Phone) ? "" : model.Phone, false);
                bookmarkNavigator.MoveToBookmark("MOBILE");
                bookmarkNavigator.ReplaceBookmarkContent(string.IsNullOrEmpty(model.Mobile) ? "" : model.Mobile, false);


                //HTML Contents: Desciprion1                
                Spire.Doc.Section tempSection = document.AddSection();
                string html = string.IsNullOrEmpty(model.Description1) ? "" : HttpUtility.HtmlDecode(model.Description1);
                tempSection.AddParagraph().AppendHTML(html);
                ParagraphBase replacementFirstItem = tempSection.Paragraphs[0].Items.FirstItem as ParagraphBase;
                ParagraphBase replacementLastItem = tempSection.Paragraphs[tempSection.Paragraphs.Count - 1].Items.LastItem as ParagraphBase;
                TextBodySelection selection = new TextBodySelection(replacementFirstItem, replacementLastItem);
                //將內容各段落套用指定的樣式
                for (int i = 0; i < tempSection.Paragraphs.Count; i++)
                {
                    tempSection.Paragraphs[i].ApplyStyle("Basic");
                }
                TextBodyPart part = new TextBodyPart(selection);

                // locate the bookmark
                bookmarkNavigator.MoveToBookmark("DESCRIPTION1");
                //replace the content of bookmark
                bookmarkNavigator.ReplaceBookmarkContent(part);
                //remove temp section
                document.Sections.Remove(tempSection);


                //HTML Contents: Desciprion2                
                tempSection = document.AddSection();
                html = string.IsNullOrEmpty(model.Description2) ? "" : HttpUtility.HtmlDecode(model.Description2);
                tempSection.AddParagraph().AppendHTML(html);
                replacementFirstItem = tempSection.Paragraphs[0].Items.FirstItem as ParagraphBase;
                replacementLastItem = tempSection.Paragraphs[tempSection.Paragraphs.Count - 1].Items.LastItem as ParagraphBase;
                selection = new TextBodySelection(replacementFirstItem, replacementLastItem);
                part = new TextBodyPart(selection);

                bookmarkNavigator.MoveToBookmark("DESCRIPTION2");
                bookmarkNavigator.ReplaceBookmarkContent(part);
                document.Sections.Remove(tempSection);


                //圖片
                bookmarkNavigator.MoveToBookmark("IMG");
                Spire.Doc.Section section_img = document.AddSection();
                Spire.Doc.Documents.Paragraph paragraph_img = section_img.AddParagraph();
                Image img = Image.FromFile(HttpContext.Current.Server.MapPath("~/App_Data/Penguins.jpg"));
                DocPicture picture = paragraph_img.AppendPicture(img);

                bookmarkNavigator.InsertParagraph(paragraph_img);
                document.Sections.Remove(section_img);

                #endregion

                #region 動態新增表格

                if (model.JobHistory.Count > 0)
                {
                    Spire.Doc.Section s = document.AddSection();
                    Spire.Doc.Table table = s.AddTable(true);
                    string[] Header = { "序號", "任職公司", "職稱", "開始時間", "結束時間" };

                    //Add Cells
                    table.ResetCells(model.JobHistory.Count + 1, Header.Length);

                    //Header Row
                    TableRow FRow = table.Rows[0];
                    FRow.IsHeader = true;
                    for (int i = 0; i < Header.Length; i++)
                    {
                        Spire.Doc.Documents.Paragraph p = FRow.Cells[i].AddParagraph();
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
                            Spire.Doc.Documents.Paragraph p2 = DataRow.Cells[c].AddParagraph();
                            TextRange TR2 = p2.AppendText(data[c]);

                            //Format Cells
                            p2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        }
                    }

                    bookmarkNavigator.MoveToBookmark("TABLE");
                    bookmarkNavigator.InsertTable(table);
                }

                #endregion

                #region 套用樣式

                //套用文章段落樣式
                for (int s = 0; s < document.Sections.Count; s++)
                {
                    Spire.Doc.Section sections = document.Sections[s];
                    //套用文章段落樣式
                    for (int p = 0; p < sections.Paragraphs.Count; p++)
                    {
                        Spire.Doc.Documents.Paragraph pgh = sections.Paragraphs[p];
                        pgh.ApplyStyle("Basic");
                        pgh.Format.BeforeSpacing = 12;
                    }

                    //套用表格樣式
                    for (int t = 0; t < document.Sections[s].Tables.Count; t++)
                    {
                        Spire.Doc.Table table = (Spire.Doc.Table)document.Sections[s].Tables[t];
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


        /// <summary>
        /// 透過既有的套印檔建立 Word 文件，並匯出 PDF 文件 (使用 Microsoft.Office.Interop.Word 套件)
        /// </summary>
        /// <param name="result">回傳: 執行結果</param>
        /// <param name="msg">回傳: 訊息</param>
        /// <returns>串流資訊</returns>
        public byte[] ExportResume_Word2PDF(out bool result, out string msg)
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

            result = true;
            msg = "";
            MemoryStream ms = new MemoryStream();
            FileStream fs = null;

            //Word 檔案套印後回傳串流
            byte[] objFile = ExportResumeByDocx_Bookmark(out result, out msg);

            //將 Word 串流資訊轉換為實體暫存檔案
            Spire.Doc.Document spiredoc = new Spire.Doc.Document();
            Stream tmpdoc = new MemoryStream(objFile);
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
                            //產生建立
                            wordDocument.ExportAsFixedFormat(tmpPdfFilePath, WdExportFormat.wdExportFormatPDF);
                            fs = new FileStream(tmpPdfFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                            wordDocument.Close();
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

            //回傳串流資訊
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
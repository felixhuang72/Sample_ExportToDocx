using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ResumeExport.Models;
using NotesFor.HtmlToOpenXml;
using System.Linq;
using System.Text;
using System.Web;

namespace ResumeExport.Service
{
    //檔案匯出服務 (DocX)
    public class OpenXmlExportService
    {
        public byte[] ExportByHtml(out bool result, out string msg)
        {
            result = true;
            msg = "";
            MemoryStream ms = new MemoryStream();

            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    // 建立 MainDocumentPart 類別物件 mainPart，加入主文件部分 
                    MainDocumentPart mainPart = doc.AddMainDocumentPart();
                    // 實例化 Document(w) 部分
                    mainPart.Document = new Document();
                    // 建立 Body 類別物件，於加入 Doucment(w) 中加入 Body 內文
                    Body body = mainPart.Document.AppendChild(
                        new Body(
                            new SectionProperties(new PageMargin()
                            {
                                Left = 600,
                                Right = 600,
                                Bottom = 700,
                                Top = 740
                            })));


                    #region 產出內容

                    // 建立 Paragraph 類別物件，於 Body 本文中加入段落 Paragraph(p)                    
                    Paragraph paragraph;
                    Run run;

                    paragraph = body.AppendChild(new Paragraph());
                    // 建立 Run 類別物件，於 段落 Paragraph(p) 中加入文字屬性 Run(r) 範圍
                    run = paragraph.AppendChild(new Run());
                    // 在文字屬性 Run(r) 範圍中加入文字內容
                    run.AppendChild(new Text("履歷匯出範例"));
                    run.AppendChild(new Break());
                    run.AppendChild(new Text("(使用 HTML 直接匯出)"));


                    //建立要產出的 HTML 內容
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
                    
                    //將 HTML 內容轉換成 XML，並添加至文件內
                    HtmlConverter converter = new HtmlConverter(mainPart);
                    converter.ParseHtml(html.ToString());
                    
                    #endregion

                    #region 套用樣式
                    
                    foreach (var p in mainPart.Document.Descendants<Paragraph>())
                    {
                        ApplyStyleToParagraph(doc, "BasicParagraphStyle", "Basic Paragraph Style", p);
                    }

                    #endregion
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


        #region 定義樣式
        //參考連結: https://msdn.microsoft.com/en-us/library/office/cc850838.aspx

        // Apply a style to a paragraph.
        private static void ApplyStyleToParagraph(WordprocessingDocument doc, string styleid, string stylename, Paragraph p)
        {
            // If the paragraph has no ParagraphProperties object, create one.
            if (p.Elements<ParagraphProperties>().Count() == 0)
            {
                p.PrependChild<ParagraphProperties>(new ParagraphProperties());
            }

            // Get the paragraph properties element of the paragraph.
            ParagraphProperties pPr = p.Elements<ParagraphProperties>().First();

            // Get the Styles part for this document.
            StyleDefinitionsPart part =
                doc.MainDocumentPart.StyleDefinitionsPart;

            // If the Styles part does not exist, add it and then add the style.
            if (part == null)
            {
                part = AddStylesPartToPackage(doc);
                AddNewStyle(part, styleid, stylename);
            }
            else
            {
                // If the style is not in the document, add it.
                if (IsStyleIdInDocument(doc, styleid) != true)
                {
                    // No match on styleid, so let's try style name.
                    string styleidFromName = GetStyleIdFromStyleName(doc, stylename);
                    if (styleidFromName == null)
                    {
                        AddNewStyle(part, styleid, stylename);
                    }
                    else
                        styleid = styleidFromName;
                }
            }

            // Set the style of the paragraph.
            pPr.ParagraphStyleId = new ParagraphStyleId() { Val = styleid };
        }

        // Return true if the style id is in the document, false otherwise.
        private static bool IsStyleIdInDocument(WordprocessingDocument doc, string styleid)
        {
            // Get access to the Styles element for this document.
            Styles s = doc.MainDocumentPart.StyleDefinitionsPart.Styles;

            // Check that there are styles and how many.
            int n = s.Elements<Style>().Count();
            if (n == 0)
                return false;

            // Look for a match on styleid.
            Style style = s.Elements<Style>()
                .Where(st => (st.StyleId == styleid) && (st.Type == StyleValues.Paragraph))
                .FirstOrDefault();
            if (style == null)
                return false;

            return true;
        }

        // Return styleid that matches the styleName, or null when there's no match.
        private static string GetStyleIdFromStyleName(WordprocessingDocument doc, string styleName)
        {
            StyleDefinitionsPart stylePart = doc.MainDocumentPart.StyleDefinitionsPart;
            string styleId = stylePart.Styles.Descendants<StyleName>()
                .Where(s => s.Val.Value.Equals(styleName) &&
                    (((Style)s.Parent).Type == StyleValues.Paragraph))
                .Select(n => ((Style)n.Parent).StyleId).FirstOrDefault();
            return styleId;
        }

        // Create a new style with the specified styleid and stylename and add it to the specified
        // style definitions part.
        private static void AddNewStyle(StyleDefinitionsPart styleDefinitionsPart, string styleid, string stylename)
        {
            // Get access to the root element of the styles part.
            Styles styles = styleDefinitionsPart.Styles;

            // Create a new paragraph style and specify some of the properties.
            Style style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = styleid,
                CustomStyle = true
            };
            StyleName styleName1 = new StyleName() { Val = stylename };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            style.Append(styleName1);
            style.Append(basedOn1);
            style.Append(nextParagraphStyle1);

            // Create the StyleRunProperties object and specify some of the run properties.
            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            Color color1 = new Color() { ThemeColor = ThemeColorValues.Accent2 };
            SpacingBetweenLines spacing = new SpacingBetweenLines() { After = "50" };

            RunFonts font1 = new RunFonts() { EastAsia = "DFKai-SB" /*正黑體: Microsoft JhengHei, 標楷體：DFKai-SB, 細明體：MingLiU, 新細明體：PMingLiU */};
            //Italic italic1 = new Italic();
            //Specify a 12 point size.
            FontSize fontSize1 = new FontSize() { Val = "24" };
            //styleRunProperties1.Append(bold1);
            //styleRunProperties1.Append(color1);
            styleRunProperties1.Append(font1);
            styleRunProperties1.Append(fontSize1);
            //styleRunProperties1.Append(italic1);
            styleRunProperties1.Append(spacing);

            // Add the run properties to the style.
            style.Append(styleRunProperties1);

            // Add the style to the styles part.
            styles.Append(style);
        }

        // Add a StylesDefinitionsPart to the document.  Returns a reference to it.
        private static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc)
        {
            StyleDefinitionsPart part;
            part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            Styles root = new Styles();
            root.Save(part);
            return part;
        }

        #endregion
    }
}
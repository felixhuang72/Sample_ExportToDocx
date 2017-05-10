<h3>專案內容</h3>
<hr />
以履歷表管理情境為例，示範如何將履歷資料以 Word 方式匯出。本專案提供兩種檔案匯出方式：

1. 完全使用 HTML 方式直接轉存匯出
2. 透過套印檔方式，將固定欄位的資料進行取代。無法事先制定格式的內容，以程式動態產生 (如: 表格內容)
<br />
<br />

<h3>使用套件</h3>
<hr />

使用 [E-ICEBLUE Free Spire.Doc for .NET](https://www.e-iceblue.com/Introduce/free-doc-component.html#.WNI-qUaP8dU)、[OpenXML SDK](https://www.nuget.org/packages/OpenXMLSDK-MOT/) 及 [Microsoft.Office.Interop.Word](https://www.nuget.org/packages/Microsoft.Office.Interop.Word/)
套件 (免費版) 進行開發<br />
Nuget 安裝指令:<br />
<pre><code>PM> Install-Package FreeSpire.Doc</code>
<code>PM> Install-Package OpenXMLSDK-MOT</code>
<code>PM> Install-Package Microsoft.Office.Interop.Word</code></pre>

<br />Spire.Doc 文件與範例：[使用文件 ](https://www.e-iceblue.com/Tutorials/Spire.Doc/Spire.Doc-Program-Guide.html)
<br />免費版限制：

- 最多 500 個段落 (Paragraph)
- 最多 25 個表格 (Table)
- 若轉存成 PDF 或 XPS 檔案，最多僅會匯出前三頁內容


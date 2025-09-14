// Author: jacques.blazor@gmail.com datetime: 2025.9.9 12:30pm License: None, all rights are reversed.
// Refactored version with separated data collection and Excel generation
// nuget package "Selenium.WebDriver" Version 4.34.0
// nuget package "EPPlus" Version 8.0.8
// Use EPPlus in a noncommercial context according to the Polyform Noncommercial license  
// ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
// dotnet publish -r win-x64 -c Release --self-contained /p:PublishSingleFile=true

using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Drawing;


ExcelPackage.License.SetNonCommercialOrganization("Noncommercial Organization");
string siteUrl = "https://www.jkforum.net/";
string parseUrl = string.Empty;

Dictionary<string, string> areaUrls = new Dictionary<string, string>
{
    {"greater_taipei", "1128-1949"},
    {"linsen_district", "1128-1476"},
    {"wanhua_ximending", "1128-1948"},
    {"taoyuan_area", "1128-1947"},
    {"banqiao_zhongyonghe", "1128-1950"},
    {"sanchong_xinzhuang", "1128-1951"},
    {"nangang_neihu", "1128-2450"},
    {"zhongli", "1128-1477"},
    {"hsinchu", "1128-1946"},
    {"taichung", "1128-1478"},
    {"beitun_district", "1128-1482"},
    {"xitun_district_fengjia", "1128-2420"},
    {"changhua_yuanlin", "1128-2421"},
    {"tainan", "1128-1479"},
    {"kaohsiung", "1128-1480"},
    {"chiayi_yunlin", "1128-2431"},
    {"yilan_huadong", "1128-1481"},
    {"other_counties_cities", "1128-1483"}
};

// 使用者選擇地區
Console.WriteLine("請選擇地區(版本:2025/9/9)：");
int index = 1;
foreach (var areaUrl in areaUrls)
{
    Console.WriteLine($"{index}. {areaUrl.Key}");
    index++;
}

string selectedArea = string.Empty;
do
{
    Console.Write("\n請輸入選擇的編號：");
    if (int.TryParse(Console.ReadLine(), out int choice) && choice >= 1 && choice <= areaUrls.Count)
    {
        selectedArea = areaUrls.Keys.ElementAt(choice - 1);
        parseUrl = $"{siteUrl}p/type-{areaUrls[selectedArea]}.html";
        Console.WriteLine($"\n選擇 {choice}.{selectedArea}。網址是：{parseUrl}。");
    }
    else
    {
        Console.WriteLine("無效的選擇。\n");
    }
} while (selectedArea == string.Empty);

// 使用者輸入頁數
Console.Write("按 Enter 接受預設為 1 頁或請輸入總頁數：");
if (!int.TryParse(Console.ReadLine(), out int capturePages))
{
    capturePages = 1;
}

try
{
    // 第一階段：擷取資料
    Console.WriteLine($"預計擷取 {capturePages} 頁。開始擷取資料，如果發生錯誤請參考記錄檔。\n擷取每頁大概需要 1~3 分鐘，請稍候...");
    var scraper = new JKForumScraper(siteUrl);
    var posts = scraper.ScrapeData(parseUrl, capturePages);

    Console.WriteLine("\n完成資料擷取。");
    // 第二階段：生成Excel
    Console.WriteLine("開始產生 Excel 檔案...");
    var excelGenerator = new ExcelGenerator();
    excelGenerator.GenerateExcel(posts, selectedArea);

    Console.WriteLine("程式執行完成！請按 Enter 結束。");
}
catch (Exception ex)
{
    Console.WriteLine($"程式執行時發生錯誤：{ex.Message}");
    System.Diagnostics.Debug.WriteLine($"錯誤詳細資訊：{ex}");
}

Console.ReadLine();

// 資料模型的類別
public class ForumPost
{
    public int Index { get; set; }              // 筆
    public string? AvatarUrl { get; set; }      // 頭像
    public string? ArticleLink { get; set; }    // 文章連結
    public string? PinToTop { get; set; }       // 置頂
    public string? AreaName { get; set; }       // 地區
    public string? AreaLink { get; set; }       // 地區連結
    public string? Availability { get; set; }   // 現在有空
    public string? Zone { get; set; }           // 分區
    public string? Title { get; set; }          // 標題
    public string? TitleLink { get; set; }      // 標題連結
    public string? Replies { get; set; }        // 回覆
    public string? Views { get; set; }          // 觀看
    public string? AuthorName { get; set; }     // 發文者
    public string? AuthorLink { get; set; }     // 發文者連結
    public string? PostDate { get; set; }       // 發文日期
    public string? LastReply { get; set; }      // 最後回覆
    public string? ReplierLink { get; set; }    // 回覆者連結
    public string? ReplyPostLink { get; set; }  // 回覆文連結
    public string? ReplyDateTime { get; set; }  // 回覆日期時間
}

// 擷取網頁的類別
public class JKForumScraper
{
    private readonly ChromeOptions _options;
    private readonly ChromeDriverService _service;
    private readonly string _siteUrl;
    private readonly string[] _areaZones;

    public JKForumScraper(string siteUrl)
    {
        _service = ChromeDriverService.CreateDefaultService();
        _service.SuppressInitialDiagnosticInformation = true; // 不要顯示啟動資訊
        _service.HideCommandPromptWindow = true;              // 隱藏黑視窗
        _service.LogPath = "JKForumScraperChromeDriver.log";  // 可選：把 log 輸出到檔案
        _service.EnableVerboseLogging = false;

        _options = new ChromeOptions();
        _options.AddArgument("--headless=new");
        _options.AddArgument("--disable-gpu");
        _options.AddArgument("--no-sandbox");
        _options.AddArgument("--silent");
        _options.AddArgument("--log-level=3");               // 0=INFO, 1=WARNING, 2=ERROR, 3=FATAL

        _siteUrl = siteUrl;
        _areaZones = new string[] { "南門", "復興路", "大興西路", "車站", "三民", "藝文", "桃園", "八德", "內壢", "鶯歌", "中壢", "蘆竹", "南崁", "大園", "楊梅", "觀音", "龜山", "長庚", "林口", "三峽", "龍潭" };
    }

    public List<ForumPost> ScrapeData(string url, int pages)
    {
        var posts = new List<ForumPost>();

        using (IWebDriver browser = new ChromeDriver(_service, _options))
        {
            browser.Navigate().GoToUrl(url);
            WebDriverWait wait = new WebDriverWait(browser, TimeSpan.FromSeconds(10));

            // 處理彈窗
            HandlePopups(browser, wait);

            // 資料擷取和多頁處理
            for (int currentPage = 1; currentPage <= pages; currentPage++)
            {
                Console.WriteLine($"\n=== 正在擷取第 {currentPage} 頁 ===");

                var pagePosts = ExtractPostsFromCurrentPage(browser, wait);
                posts.AddRange(pagePosts);

                Console.WriteLine($"\n第 {currentPage} 頁資料擷取完成，實際取得 {pagePosts.Count}/{posts.Count} 筆貼文。");

                // 換頁
                if (currentPage < pages && !GoToNextPage(browser, wait))
                {
                    Console.WriteLine("沒有更多頁面了");
                    break;
                }
            }

            browser.Quit();
        }

        // 為貼文增加索引
        for (int i = 0; i < posts.Count; i++)
        {
            // 從第2行開始，因為第1行是標題，筆數從 1 開始算
            posts[i].Index = i + 1;
        }

        return posts;
    }

    private void HandlePopups(IWebDriver browser, WebDriverWait wait)
    {
        try
        {
            wait.Until(driver => driver.FindElement(By.Id("fd_page_bottom")));
        }
        catch
        {
            System.Diagnostics.Debug.WriteLine("當等待頁面載入時逾時！");
        }

        try
        {
            var dontRemind = wait.Until(driver => driver.FindElement(By.XPath("//*[@id='periodaggre18']")));
            dontRemind.Click();
        }
        catch
        {
            System.Diagnostics.Debug.WriteLine("在處理 periodaggre18 時發生彈窗錯誤！");
        }

        try
        {
            var yesOver18 = wait.Until(driver => driver.FindElement(By.XPath("//*[@id='fwin_dialog_submit']")));
            yesOver18.Click();
        }
        catch
        {
            System.Diagnostics.Debug.WriteLine("在處理 fwin_dialog_submit 時發生彈窗錯誤！");
        }

        try
        {
            wait.Until(driver => driver.FindElement(By.ClassName("nxt")));
        }
        catch
        {
            System.Diagnostics.Debug.WriteLine("在等待頁面準備完成時逾時！");
        }
    }

    private List<ForumPost> ExtractPostsFromCurrentPage(IWebDriver browser, WebDriverWait wait)
    {
        var posts = new List<ForumPost>();

        try
        {
            // 取得 HTML 表格
            var table = wait.Until(driver => driver.FindElement(By.Id("threadlisttableid")));
            // 分別從 HTML 表格裡的每一列 row 裡旳 tr 
            var trs = table.FindElements(By.TagName("tr"));

            Console.Write($"開始擷取表格資料，預計有 {trs.Count} 筆資料 ");

            foreach (var tr in trs)
            {
                Console.Write($".");
                // 分別從每一列 tr 中擷取出每一筆貼文的內容
                var post = ExtractPostFromRow(tr);
                if (post != null)
                {
                    posts.Add(post);
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"在擷取頁面資料時發生錯誤：{ex.Message}！");
        }

        return posts;
    }

    // 取出貼文
    private ForumPost? ExtractPostFromRow(IWebElement tr)
    {
        try
        {
            // 貼文
            var post = new ForumPost();

            // 處理 th 標籤 (標題行)
            System.Diagnostics.Debug.WriteLine($"== 擷取標題 ==");
            // 標題
            var ths = tr.FindElements(By.TagName("th"));
            if (ths.Count == 1)
            {
                var trParentId = tr.FindElement(By.XPath("..")).GetAttribute("id");
                if (trParentId == "separatorline_top")
                {
                    return null;
                }
                ExtractHeaderData(ths[0], post);
            }

            // 處理從 tr 中擷取出資料列中 td 標籤的內容
            System.Diagnostics.Debug.WriteLine($"== 擷取每筆貼文 ==");
            var tds = tr.FindElements(By.TagName("td"));
            if (tds.Count == 3)
            {
                // 發文者、發文者連結、發文日期、最後回覆、回覆者連結、回覆文連結 和 回覆日期時間
                ExtractCellData(tds, post);
            }

            // 只有當貼文有標題時才回傳
            return !string.IsNullOrEmpty(post.Title) ? post : null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"在擷取頁面貼文資料時發生資料錯誤: {ex.Message}");
            return null;
        }
    }

    // 標題資料
    private void ExtractHeaderData(IWebElement th, ForumPost post)
    {
        try
        {
            // 擷取頭像資訊
            System.Diagnostics.Debug.WriteLine($"== 擷取頭像 ==");
            var thSpanA = th.FindElement(By.TagName("a"));
            if (thSpanA != null)
            {
                // 頭像影像
                var thumbnailImage = thSpanA.FindElement(By.TagName("img"));
                if (thumbnailImage != null)
                {
                    // 取得頭像影像檔連結
                    post.AvatarUrl = thumbnailImage.GetAttribute("src");
                    // 取得文章連結
                    post.ArticleLink = thSpanA.GetAttribute("href");
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"在擷取頭像及文章連結資料時發生資料錯誤: {ex.Message}");
        }

        // 擷取div資訊
        System.Diagnostics.Debug.WriteLine($"== 擷取div資訊 ==");
        var divs = th.FindElements(By.TagName("div"));
        if (divs.Count == 3)
        {
            ExtractDivData(divs, post);
        }
    }

    // 置頂(pin), 回覆(posted), 觀看(viewed) 
    private void ExtractDivData(IList<IWebElement> divs, ForumPost post)
    {
        System.Diagnostics.Debug.WriteLine($"== 擷取div資料 ==");
        for (int i = 0; i < divs.Count; i++)
        {
            switch (i)
            {
                // 置頂
                case 0:
                    ExtractFirstDivData(divs[i], post);
                    break;
                // case 1 這部份不需要處理直接跳過
                case 1:
                    // 這部份不需要處理直接跳過
                    break;
                // 處理 回覆(posted) 和 觀看(viewed) 的值
                case 2:
                    ExtractThirdDivData(divs[i], post);
                    break;
            }
        }
    }

    // 置頂
    private void ExtractFirstDivData(IWebElement div, ForumPost post)
    {
        // 置頂資訊
        System.Diagnostics.Debug.WriteLine($"== 擷取置頂資訊 ==");
        try
        {
            var pinImg = div.FindElement(By.TagName("img"));
            string pinToTop = _siteUrl + pinImg.GetAttribute("src");
            post.PinToTop = string.Empty;
            if (pinToTop == "https://www.jkforum.net/https://www.jkforum.net/template/yibai_city1/style/pin_2.gif")
            {
                post.PinToTop = "置頂";
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"在擷取置頂資訊時發生錯誤：{ex.Message}！");
        }

        // 地區資訊
        System.Diagnostics.Debug.WriteLine($"== 擷取地區資訊 ==");
        try
        {
            var areaLink = div.FindElement(By.TagName("a"));
            // 地區
            post.AreaName = areaLink.Text;
            // 地區連結
            post.AreaLink = areaLink.GetAttribute("href");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"在擷取地區資訊時發生錯誤：{ex.Message}！");
        }

        // 現在有空資訊
        System.Diagnostics.Debug.WriteLine($"== 擷取現在有空資訊 ==");
        try
        {
            var divImages = div.FindElements(By.TagName("img"));
            if (divImages.Count == 2)
            {
                // 現在有空
                string availabilityUrl = _siteUrl + divImages[1].GetAttribute("src");
                post.Availability = string.Empty;
                if (availabilityUrl == "https://www.jkforum.net/https://www.jkforum.net/static/image/common/freenow.png")
                {
                    post.Availability = "有空";
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"在擷取現在有空資訊時發生錯誤：{ex.Message}！");
        }

        // 標題、標題連結、分區
        System.Diagnostics.Debug.WriteLine($"== 擷取標題和區域資訊 ==");
        try
        {
            var titleLinks = div.FindElements(By.CssSelector("a.s.xst"));
            if (titleLinks.Count == 1 && !string.IsNullOrEmpty(titleLinks[0].Text))
            {
                // 標題
                var title = titleLinks[0].Text;
                // 移除開頭的 '+' 或 '=' 特殊字元
                title = title.TrimStart('+', '=');
                // 標題
                post.Title = title;
                // 標題連結
                post.TitleLink = titleLinks[0].GetAttribute("href");

                // 分區
                post.Zone = _areaZones.FirstOrDefault(zone => title.Contains(zone));
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"在擷取標題和區域資訊時發生錯誤：{ex.Message}！");
        }
    }

    // 回覆(posted)、觀看(viewed) 
    private void ExtractThirdDivData(IWebElement div, ForumPost post)
    {
        System.Diagnostics.Debug.WriteLine($"== 擷取 ThirdDivData ==");
        // 移除空白字元
        var divText = div.Text?.Replace(" ", "").Replace("\t", "").Replace("\n", "").TrimEnd('/');

        // 判斷是否只包含一個 '/'
        if (!string.IsNullOrEmpty(divText) && divText.Count(f => f == '/') == 1)
        {
            var parts = divText.Split('/');
            // 回覆
            post.Replies = parts[0];
            // 觀看
            post.Views = parts[1];
        }
    }

    // 發文者、發文者連結、發文日期、最後回覆、回覆者連結、回覆文連結 和 回覆日期時間
    private void ExtractCellData(IList<IWebElement> tds, ForumPost post)
    {
        System.Diagnostics.Debug.WriteLine($"== 擷取 CellData ==");
        for (int i = 0; i < tds.Count; i++)
        {
            switch (i)
            {
                // 發文者、發文者連結和發文日期
                case 0:
                    ExtractAuthorData(tds[i], post);
                    break;
                // 這部份不需要處理直接跳過
                case 1:
                    break;
                // 最後回覆
                case 2:
                    ExtractReplyData(tds[i], post);
                    break;
            }
        }
    }

    // 發文者、發文者連結和發文日期
    private void ExtractAuthorData(IWebElement td, ForumPost post)
    {
        System.Diagnostics.Debug.WriteLine($"== 擷取 AuthorData ==");
        try
        {
            var authorLink = td.FindElement(By.TagName("cite")).FindElement(By.TagName("a"));
            // 發文者
            post.AuthorName = authorLink.Text;
            // 發文者連結
            post.AuthorLink = authorLink.GetAttribute("href");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"在擷取 Author 資時發生錯誤：{ex.Message}！");
        }
        try
        {
            var dateSpan = td.FindElement(By.TagName("em")).FindElement(By.TagName("span"));
            // 發文日期
            post.PostDate = dateSpan.Text;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"在擷取 PostDate 資料時發生錯誤：{ex.Message}！");
        }
    }

    // 最後回覆
    private void ExtractReplyData(IWebElement td, ForumPost post)
    {
        System.Diagnostics.Debug.WriteLine($"== 擷取 LastReply ==");
        try
        {
            var replyLink = td.FindElement(By.TagName("cite")).FindElement(By.TagName("a"));
            // 最後回覆
            post.LastReply = replyLink.Text;
            // 回覆者連結
            post.ReplierLink = _siteUrl + replyLink.GetAttribute("href");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"在擷取 LastReply 資料時發生錯誤：{ex.Message}！");
        }
        System.Diagnostics.Debug.WriteLine($"== 擷取 replyPostLink ==");
        try
        {
            var replyPostLink = td.FindElement(By.TagName("em")).FindElement(By.TagName("a"));
            // 回覆文連結
            post.ReplyPostLink = replyPostLink.GetAttribute("href");
            // 回覆日期時間
            post.ReplyDateTime = replyPostLink.Text.Replace("\xa0", "");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"在擷取 replyPostLink 資料時發生錯誤：{ex.Message}！");
        }
    }

    // 換到下一頁
    private bool GoToNextPage(IWebDriver browser, WebDriverWait wait)
    {
        System.Diagnostics.Debug.WriteLine($"== 進行 NextPage ==");
        try
        {
            var nextButton = browser.FindElement(By.ClassName("nxt"));
            nextButton.Click();
            wait.Until(driver => driver.FindElement(By.Id("fd_page_bottom")));
            return true;
        }
        catch (NoSuchElementException)
        {
            System.Diagnostics.Debug.WriteLine("已經擷取完了沒有其他資料！");
            return false;
        }
        catch (TimeoutException)
        {
            System.Diagnostics.Debug.WriteLine("在等待下一個頁面載入時逾時！");
            return false;
        }
    }
}

// 產生Excel的類別
public class ExcelGenerator
{
    private readonly string[] _headers = new string[]
    {
        "筆", "頭像", "文章連結", "置頂", "地區", "地區連結", "現在有空", "分區",
        "標題", "標題連結", "回覆", "觀看", "發文者", "發文者連結", "發文日期",
        "最後回覆", "回覆者連結", "回覆文連結", "回覆日期時間"
    };

    private readonly int[] _columnWidths = new int[]
    {
        4, 8, 8, 4, 4, 7, 7, 8, 72, 8, 5, 14, 10, 11, 13, 8, 15, 15, 18
    };

    public void GenerateExcel(List<ForumPost> posts, string selectedArea)
    {
        System.Diagnostics.Debug.WriteLine($"== 進行 GenerateExcel ==");
        string filePath = GetUniqueFilePath(selectedArea);

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets.Add($"JKF{selectedArea}");

            SetupStyles(package);
            WriteHeaders(worksheet);
            WriteData(worksheet, posts);
            FormatWorksheet(worksheet);

            package.Save();

            Console.WriteLine($"檔案 {filePath} 已寫入完成，請按 Enter 鍵結束。");
        }
    }

    private string GetUniqueFilePath(string selectedArea)
    {
        System.Diagnostics.Debug.WriteLine($"== 進行 GetUniqueFilePath ==");
        string fileName;
        string filePath;

        do
        {
            string currentDate = DateTime.Now.ToString("yyyyMMddHHmmss");
            fileName = $"JKF{selectedArea}{currentDate}.xlsx";
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            filePath = Path.Combine(desktopPath, fileName);
        }
        while (File.Exists(filePath));

        return filePath;
    }

    private void SetupStyles(ExcelPackage package)
    {
        System.Diagnostics.Debug.WriteLine($"== 進行 SetupStyles ==");
        var hyperLinkStyle = package.Workbook.Styles.CreateNamedStyle("HyperLink");
        hyperLinkStyle.Style.Font.UnderLine = true;
        hyperLinkStyle.Style.Font.Color.SetColor(Color.Blue);
    }

    private void WriteHeaders(ExcelWorksheet worksheet)
    {
        System.Diagnostics.Debug.WriteLine($"== 進行 WriteHeaders ==");
        for (int i = 0; i < _headers.Length; i++)
        {
            worksheet.Cells[1, i + 1].Value = _headers[i];
        }

        char endColumn = (char)('A' + _headers.Length - 1);
        using (var titleRange = worksheet.Cells[$"A1:{endColumn}1"])
        {
            titleRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            titleRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(128, 96, 0));
            titleRange.Style.Font.Color.SetColor(Color.White);
            titleRange.Style.Font.Bold = true;
            titleRange.AutoFilter = true;
        }
    }

    private void WriteData(ExcelWorksheet worksheet, List<ForumPost> posts)
    {
        System.Diagnostics.Debug.WriteLine($"== 進行 WriteData ==");
        char endColumn = (char)('A' + _headers.Length - 1);

        Console.WriteLine($"正在寫入資料...");
        for (int i = 0; i < posts.Count; i++)
        {
            // 從第 2 列開始
            int row = i + 2;
            var post = posts[i];

            WritePostData(worksheet, row, post);

            // 設定雙數列背景色
            if (row % 2 == 1)
            {
                var range = worksheet.Cells[$"A{row}:{endColumn}{row}"];
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 242, 204));
            }
        }
        Console.WriteLine($"\n寫入資料完成！");
    }

    private void WritePostData(ExcelWorksheet worksheet, int row, ForumPost post)
    {
        // 基本資料
        worksheet.Cells[row, GetColumnIndex("筆")].Value = post.Index;
        worksheet.Cells[row, GetColumnIndex("頭像")].Value = post.AvatarUrl;
        worksheet.Cells[row, GetColumnIndex("文章連結")].Value = post.ArticleLink;
        worksheet.Cells[row, GetColumnIndex("置頂")].Value = post.PinToTop;
        worksheet.Cells[row, GetColumnIndex("地區")].Value = post.AreaName;
        worksheet.Cells[row, GetColumnIndex("地區連結")].Value = post.AreaLink;
        worksheet.Cells[row, GetColumnIndex("現在有空")].Value = post.Availability;
        worksheet.Cells[row, GetColumnIndex("分區")].Value = post.Zone;
        worksheet.Cells[row, GetColumnIndex("標題")].Value = post.Title;
        worksheet.Cells[row, GetColumnIndex("標題連結")].Value = post.TitleLink;
        worksheet.Cells[row, GetColumnIndex("發文者")].Value = post.AuthorName;
        worksheet.Cells[row, GetColumnIndex("發文者連結")].Value = post.AuthorLink;
        worksheet.Cells[row, GetColumnIndex("發文日期")].Value = post.PostDate;
        worksheet.Cells[row, GetColumnIndex("最後回覆")].Value = post.LastReply;
        worksheet.Cells[row, GetColumnIndex("回覆者連結")].Value = post.ReplierLink;
        worksheet.Cells[row, GetColumnIndex("回覆文連結")].Value = post.ReplyPostLink;
        worksheet.Cells[row, GetColumnIndex("回覆日期時間")].Value = post.ReplyDateTime;

        // 處理數字欄位
        SetNumericValue(worksheet, row, GetColumnIndex("回覆"), post.Replies);
        SetNumericValue(worksheet, row, GetColumnIndex("觀看"), post.Views);

        // 設定超連結
        SetHyperlink(worksheet, row, GetColumnIndex("頭像"), post.AvatarUrl);
        SetHyperlink(worksheet, row, GetColumnIndex("文章連結"), post.ArticleLink);
        SetHyperlink(worksheet, row, GetColumnIndex("標題"), post.TitleLink);
        SetHyperlink(worksheet, row, GetColumnIndex("發文者"), post.AuthorLink);
        SetHyperlink(worksheet, row, GetColumnIndex("最後回覆"), post.ReplierLink);
        SetHyperlink(worksheet, row, GetColumnIndex("回覆文連結"), post.ReplyPostLink);
    }

    private void SetNumericValue(ExcelWorksheet worksheet, int row, int column, string? value)
    {
        if (!string.IsNullOrEmpty(value) && int.TryParse(value, out int numericValue))
        {
            worksheet.Cells[row, column].Value = numericValue;
        }
        else
        {
            worksheet.Cells[row, column].Value = value;
        }
    }

    private void SetHyperlink(ExcelWorksheet worksheet, int row, int column, string? url)
    {
        if (!string.IsNullOrEmpty(url))
        {
            worksheet.Cells[row, column].Hyperlink = new ExcelHyperLink(url);
            worksheet.Cells[row, column].StyleName = "HyperLink";
        }
    }

    private int GetColumnIndex(string headerName)
    {
        return Array.IndexOf(_headers, headerName) + 1;
    }

    private void FormatWorksheet(ExcelWorksheet worksheet)
    {
        // 設定字型和樣式
        worksheet.Cells.Style.Font.Name = "Yu Gothic";
        worksheet.Cells.Style.Font.Size = 10;
        worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        // 設定列寬
        for (int i = 0; i < _columnWidths.Length; i++)
        {
            worksheet.Column(i + 1).Width = _columnWidths[i];
        }
    }
}
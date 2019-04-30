//#define XPATH
//#define Regex
//#define CSS

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using HtmlAgilityPack;
using Ivony.Html;

namespace Crawler.Models
{
    public static class BookHelper
    {
        #region field&property

        /// <summary>
        /// 待爬取条数
        /// </summary>
        private const int BookCount = 1000;

        /// <summary>
        /// 定义泛型列表存放书籍信息
        /// </summary>
        private static readonly List<Book> bookList = new List<Book>();

        /// <summary>
        /// 定义爬虫入口URL
        /// </summary>
        private const string BookUrl = "http://www.yousuu.com/category/all?sort=rate&page=";

        /// <summary>
        /// 当前爬取页面
        /// </summary>
        private static int _currentPageNum = 1;

        /// <summary>
        /// 错误次数
        /// </summary>
        private static int _errorCount = 0;

        #endregion

        #region public method

        public static void DoWork()
        {
            var watch = new Stopwatch();
            watch.Start();
            Console.WriteLine("任务开始...");

            while(bookList.Count < BookCount)
            {
                BookCrawler(_currentPageNum);
            }

            Console.WriteLine($"开始保存数据至excel，共{bookList.Count}条数据...");
            SaveDataToExcelFile();
            Console.WriteLine("保存完毕！");
            watch.Stop();
            Console.WriteLine($"任务结束，耗时：{watch.ElapsedMilliseconds}毫秒,出错次数{_errorCount}");
        }

        #endregion

        #region private method

        /// <summary>
        /// 抓取书籍信息
        /// </summary>
        /// <param name="currentPageNum">当前爬取页面</param>
        private static void BookCrawler(int currentPageNum)
        {
            var bookCrawler = new SimpleCrawler();
            bookCrawler.OnStart += (s, e) =>
            {
                Console.WriteLine($"开始抓取第{currentPageNum}页数据，页面地址：{ e.Uri.ToString()}");
            };
            bookCrawler.OnError += (s, e) =>
            {
                Console.WriteLine($"抓取第{currentPageNum}页数据出现错误：{e.Exception.Message},准备重新抓取...");
                _errorCount++;
            };
            bookCrawler.OnCompleted += (s, e) =>
            {
                int count = 0;
#if Regex
                ParseDataWithRegex(e.PageSource, ref count);
#elif CSS
                ParseDataWithCss(e.PageSource, ref count);
#else
                ParseDataWithXpath(e.PageSource, ref count);
#endif

                Console.WriteLine("===============================================");
                Console.WriteLine($"第{currentPageNum++}页数据抓取完成！本页合计{count}本书,当前合计{bookList.Count}本书");
                Console.WriteLine($"耗时：{e.Milliseconds}毫秒");
                Console.WriteLine($"线程：{e.ThreadId}");
                Console.WriteLine($"地址：{e.Uri.ToString()}");
                _currentPageNum++;
            };
            bookCrawler.Start(new Uri(BookUrl + currentPageNum)).Wait();
        }

        /// <summary>
        /// 使用正则表达式解析数据
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="count"></param>
        private static void ParseDataWithRegex(IHtmlDocument doc, ref int count)
        {
            // 使用正则表达式清洗网页源代码中的数据
            string pattern =
                "<div class=\"title\"><a href=\"(?<href>/book/[\\d]+)\" target=\"_blank\">(?<title>[\\w\\s《》—]+)</a></div><div class=\"rating\"><span class=\"allstar00\"></span><span class=\"rating_nums\"></span><span>\\((?<ratenumber>\\w+)评价\\)</span></div><div class=\"abstract\">作者:\\s*(?<author>[\\w\\s-]+)\\s*<br />字数:\\s*(?<wordnumber>[\\w\\.]+)\\s*<br />最后更新:\\s*(?<updatetime>[\\w\\s]+)<br />综合评分:\\s*<span class=\"num2star\">(?<score>[\\w\\.]+)</span></div>";

            var links = Regex.Matches(doc.InnerHtml().ToString(), pattern, RegexOptions.IgnoreCase);

            foreach(Match match in links)
            {
                var book = new Book
                {
                    Title = match.Groups["title"].Value,
                    Author = match.Groups["author"].Value,
                    WordNumber = match.Groups["wordnumber"].Value,
                    UpdateTime = match.Groups["updatetime"].Value,
                    Score = match.Groups["score"].Value,
                    RateNumber = match.Groups["ratenumber"].Value,
                    Url = new Uri("http://www.yousuu.com" + match.Groups["href"].Value)
                };
                lock(bookList)
                {
                    if(!bookList.Contains(book))
                    {
                        count++;
                        bookList.Add(book); //将数据加入到泛型列表
                    }
                }

                //Console.WriteLine(book.ToString());//将书籍信息显示到控制台
            }
        }

        /// <summary>
        /// 使用CSS选择器解析数据
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="count"></param>
        private static void ParseDataWithCss(IHtmlDocument doc, ref int count)
        {
            IEnumerable<IHtmlElement> books = doc.Find("body > div.sokk-body > div > div > div.col-sm-9.col-md-10.col-lg-10.col-xs-12 > div > div.books > div > div");

            foreach(IHtmlElement bookInfo in books)
            {
                var book = new Book();
                foreach(IHtmlElement info in bookInfo.Find("div > div > div > div.title >a"))
                {
                    book.Title = info.InnerText();
                    book.Url = new Uri("http://www.yousuu.com" + info.Attribute("href").Value());
                    break;
                }

                foreach(IHtmlElement info in bookInfo.Find("div > div > div > div.rating > span:nth-child(3)"))
                {
                    book.RateNumber = info.InnerText();
                }

                foreach(IHtmlElement info in bookInfo.Find("div > div > div > div.abstract"))
                {
                    string text = info.ToString();
                    var li = text.Split(':', '：', '>', '<');
                    if(li.Length != 19)
                    {
                        book = null;
                        break;
                    }

                    book.Author = li[3];
                    book.WordNumber = li[6];
                    book.UpdateTime = li[9];
                    book.Score = li[14];
                }

                lock(bookList)
                {
                    if(book != null && !bookList.Contains(book))
                    {
                        count++;
                        bookList.Add(book);//将数据加入到泛型列表
                    }
                }
                //Console.WriteLine(book?.ToString());//将书籍信息显示到控制台
            }
        }

        /// <summary>
        /// 使用Xpath解析数据
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="count"></param>
        private static void ParseDataWithXpath(IHtmlDocument doc, ref int count)
        {
            HtmlDocument hdoc = new HtmlDocument();
            hdoc.LoadHtml(doc.InnerHtml().ToString());

            HtmlNode books = hdoc.DocumentNode.SelectSingleNode("/html/body/div[1]/div/div/div[2]/div/div[3]/div");
            foreach(var bookinfo in books.ChildNodes)
            {
                var book = new Book
                {
                    Title = bookinfo.SelectSingleNode("div/div/div/div[3]/a").InnerText,
                    Author = bookinfo.SelectSingleNode("div/div/div/div[5]/text()[1]").InnerText.Replace("作者: ", ""),
                    WordNumber = bookinfo.SelectSingleNode("div/div/div/div[5]/text()[2]").InnerText,
                    UpdateTime = bookinfo.SelectSingleNode("div/div/div/div[5]/text()[3]").InnerText,
                    Score = bookinfo.SelectSingleNode("div/div/div/div[5]/span").InnerText,
                    RateNumber = bookinfo.SelectSingleNode("div/div/div/div[4]/span[3]").InnerText,
                    Url = new Uri("http://www.yousuu.com" + bookinfo.SelectSingleNode("div/div/div/div[3]/a").Attributes["href"].Value)
                };
                lock(bookList)
                {
                    if(!bookList.Contains(book))
                    {
                        count++;
                        bookList.Add(book); //将数据加入到泛型列表
                    }
                }
                //Console.WriteLine(book.ToString());//将书籍信息显示到控制台
            }
        }

        /// 将数据保存为Excel文件(指定列类型)
        /// <returns></returns>
        private static void SaveDataToExcelFile()
        {
            string filename = AppDomain.CurrentDomain.BaseDirectory + $"\\优书网TOP{BookCount }-{DateTime.Now:yyyyMMdd}.xls";
            var excelDoc = new StreamWriter(filename);
            const string startExcelXml = @"<xml version> <Workbook " +
                  "xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" " +
                  " xmlns:o=\"urn:schemas-microsoft-com:office:office\"  " +
                  "xmlns:x=\"urn:schemas-    microsoft-com:office:" +
                  "excel\"  xmlns:ss=\"urn:schemas-microsoft-com:" +
                  "office:spreadsheet\">  <Styles>  " +
                  "<Style ss:ID=\"Default\" ss:Name=\"Normal\">  " +
                  "<Alignment ss:Vertical=\"Bottom\"/>  <Borders/>" +
                  "  <Font/>  <Interior/>  <NumberFormat/>" +
                  "  <Protection/>  </Style>  " +
                  "<Style ss:ID=\"BoldColumn\">  <Font " +
                  "x:Family=\"宋体\" ss:Bold=\"0\"/>  </Style>  " +
                  "<Style     ss:ID=\"StringLiteral\">  <NumberFormat" +
                  " ss:Format=\"@\"/>  </Style>  <Style " +
                  "ss:ID=\"Decimal\">  <NumberFormat " +
                  "ss:Format=\"0.00000\"/>  </Style>  " +
                  "<Style ss:ID=\"Integer\">  <NumberFormat " +
                  "ss:Format=\"0\"/>  </Style>  <Style " +
                  "ss:ID=\"DateLiteral\">  <NumberFormat " +
                  "ss:Format=\"yyyy-MM-dd HH:mm:ss\"/>  </Style>  " +
                  "</Styles>  ";
            const string endExcelXml = "</Workbook>";

            int rowCount = 0;
            int sheetCount = 1;

            excelDoc.Write(startExcelXml);
            excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
            excelDoc.Write("<Table ss:DefaultColumnWidth=\"90\">");

            excelDoc.Write("<Row>");
            excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">书籍名称</Data></Cell>");
            excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">作者</Data></Cell>");
            excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">字数</Data></Cell>");
            excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">最后更新时间</Data></Cell>");
            excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">评分</Data></Cell>");
            excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">评价人数</Data></Cell>");
            excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">优书网链接</Data></Cell>");
            excelDoc.Write("</Row>");

            for(int i = 0; i < bookList.Count; i++)
            {
                rowCount++;
                //if the number of rows is > 64000 create a new page to continue output
                //分页
                if(rowCount % 10000 == 0)
                {
                    rowCount = 0;
                    sheetCount++;
                    excelDoc.Write("</Table>");
                    excelDoc.Write(" </Worksheet>");
                    excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
                    excelDoc.Write("<Table ss:DefaultColumnWidth=\"90\">");

                    excelDoc.Write("<Row>");
                    excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">书籍名称</Data></Cell>");
                    excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">作者</Data></Cell>");
                    excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">字数</Data></Cell>");
                    excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">最后更新时间</Data></Cell>");
                    excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">评分</Data></Cell>");
                    excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">评价人数</Data></Cell>");
                    excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">优书网链接</Data></Cell>");
                    excelDoc.Write("</Row>");
                }

                excelDoc.Write("<Row>");
                {
                    var book = bookList[i];
                    excelDoc.Write($"<Cell><Data ss:Type=\"String\">{book.Title}</Data></Cell>");
                    excelDoc.Write($"<Cell><Data ss:Type=\"String\">{book.Author}</Data></Cell>");
                    excelDoc.Write($"<Cell><Data ss:Type=\"String\">{book.WordNumber}</Data></Cell>");
                    excelDoc.Write($"<Cell><Data ss:Type=\"String\">{book.UpdateTime}</Data></Cell>");
                    excelDoc.Write($"<Cell><Data ss:Type=\"String\">{book.Score}</Data></Cell>");
                    excelDoc.Write($"<Cell><Data ss:Type=\"String\">{book.RateNumber}</Data></Cell>");
                    excelDoc.Write($"<Cell><Data ss:Type=\"String\">{book.Url}</Data></Cell>");
                }
                excelDoc.Write("</Row>");
            }

            excelDoc.Write("</Table>");
            excelDoc.Write(" </Worksheet>");
            excelDoc.Write(endExcelXml);
            excelDoc.Close();
        }

        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Crawler.Events;
using Ivony.Html;
using Ivony.Html.Parser;

namespace Crawler
{
    public class SimpleCrawler : ICrawler
    {
        public event EventHandler<OnStartEventArgs> OnStart;//爬虫启动事件

        public event EventHandler<OnCompletedEventArgs> OnCompleted;//爬虫完成事件

        public event EventHandler<OnErrorEventArgs> OnError;//爬虫出错事件

        public CookieContainer CookiesContainer { get; set; }//定义Cookie容器

        /// <summary>
        /// 异步创建爬虫
        /// </summary>
        /// <param name="uri">爬虫URL地址</param>
        /// <returns>网页源代码</returns>
        public async Task Start(Uri uri)
        {
            await Task.Run(() =>
           {
               try
               {
                   this.OnStart?.Invoke(this, new OnStartEventArgs(uri));

                   var watch = new Stopwatch();
                   watch.Start();
                   IHtmlDocument pageSource = new JumonyParser().LoadDocument(uri.AbsoluteUri);
                   watch.Stop();

                   var threadId = System.Threading.Thread.CurrentThread.ManagedThreadId;//获取当前任务线程ID
                    var milliseconds = watch.ElapsedMilliseconds;//获取请求执行时间
                    this.OnCompleted?.Invoke(this, new OnCompletedEventArgs(uri, threadId, milliseconds, pageSource));
               }
               catch(Exception ex)
               {
                   OnError?.Invoke(this, new OnErrorEventArgs(uri, ex));
               }
           });
        }
    }
}

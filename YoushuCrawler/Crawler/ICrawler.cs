using System;
using System.Threading.Tasks;
using Crawler.Events;

namespace Crawler
{
    public interface ICrawler
    {
        event EventHandler<OnStartEventArgs> OnStart;//爬虫启动事件

        event EventHandler<OnCompletedEventArgs> OnCompleted;//爬虫完成事件

        event EventHandler<OnErrorEventArgs> OnError;//爬虫出错事件

        Task Start(Uri uri); //异步爬虫
    }
}

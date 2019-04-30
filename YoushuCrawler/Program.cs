using System;
using Crawler.Models;

namespace Crawler
{
    static class Program
    {
        static void Main(string[] args)
        {
            BookHelper.DoWork();
            Console.ReadKey();
        }
    }
}



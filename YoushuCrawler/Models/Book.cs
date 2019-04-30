namespace Crawler.Models
{
    /// <summary>
    /// 书籍信息
    /// </summary>
    public class Book
    {
        /// <summary>
        /// 书籍名称
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// 作者
        /// </summary>
        public string Author { get; set; }
        /// <summary>
        /// 字数
        /// </summary>
        public string WordNumber { get; set; }
        /// <summary>
        /// 最后更新时间
        /// </summary>
        public string UpdateTime { get; set; }
        /// <summary>
        /// 评分
        /// </summary>
        public string Score { get; set; }
        /// <summary>
        /// 评价人数
        /// </summary>
        public string RateNumber { get; set; }
        /// <summary>
        /// 优书网链接
        /// </summary>
        public System.Uri Url { get; set; }

        public override string ToString()
        {
            return $"书名：{Title}|作者：{Author}|字数：{WordNumber}|最后更新时间：{UpdateTime}|评分：{Score}|评价人数：{RateNumber}|优书网链接：{Url.ToString()}";
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Hotel.VO
{
    public class Bsgrid<T>
    {
        //泛型：T
        /// <summary>
        /// 成功否
        /// </summary>
        public bool success { get; set; }
        /// <summary>
        /// 总的行数
        /// </summary>
        public int totalRows { get; set; }
        /// <summary>
        /// 当前页
        /// </summary>
        public int curPage { get; set; }
        /// <summary>
        /// 数据：泛型数据集
        /// </summary>
        public List<T> data { get; set; }
    }
}
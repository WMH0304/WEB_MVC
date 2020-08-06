﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Hotel.VO
{
    public class BsgridPage
    {
        /// <summary>
        /// 每一页数据的数目
        /// </summary>
        public int pageSize { get; set; }
        /// <summary>
        /// 当前页
        /// </summary>
        public int curPage { get; set; }

        public string sortName { get; set; }

        public string sortOrder { get; set; }
        /// <summary>
        /// 当前页开始序号
        /// </summary>
        public int GetStartIndex()
        {
            return (curPage - 1) * pageSize;
        }
        /// <summary>
        /// 当前页结束序号
        /// </summary>
        public int GetEndIndex()
        {
            return curPage * pageSize - 1;
        }
    }
}
using SeaTransportation.VO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SeaTransportation.Common
{
    public class Tools
    {
        public static List<SelectVo> SetSelectJson(List<SelectVo> select)
        {
            List<SelectVo> list = new List<SelectVo>();
            SelectVo selectVo = new SelectVo
            {
                id = 0,
                text = "——请选择——"
            };
            list.Add(selectVo);
            list.AddRange(select);
            return list;
        }
    }
}
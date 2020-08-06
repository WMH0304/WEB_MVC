using Hotel.VO;
using SeaTransportation.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Web.Mvc;

namespace SeaTransportation.Areas.PPAP.Controllers
{
    public class SystemAAMController : Controller
    {
        Models.SeaTransportationEntities myModels = new Models.SeaTransportationEntities();
        // GET: PPAP/SystemAAM
        /// <summary>
        /// 组织结构
        /// </summary>
        /// <returns></returns>
        public ActionResult TissueStructure()
        {
            return View();
        }

        #region 组织结构代码

        /// <summary>
        /// 树形
        /// </summary>
        /// <returns></returns>
        public ActionResult Tree()
        {
            var list = (from tbMessage in myModels.SYS_Message
                        select new
                        {
                            id = tbMessage.MessageID,
                            name = tbMessage.ChineseName
                        });
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public ActionResult rrr()
        {
            var ggg = (from tbDepartment in myModels.SYS_Department
                       select new
                       {
                           id = tbDepartment.DepartmentID,
                           name = tbDepartment.DepartmentName
                       });
            return Json(ggg, JsonRequestBehavior.AllowGet);
        }
        public ActionResult www()
        {
            var bbb = (from tbJurisdiction in myModels.SYS_Jurisdiction select new { id = tbJurisdiction.JurisdictionID, name = tbJurisdiction.JurisdictionName });
            return Json(bbb, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询
        /// </summary>
        /// <returns></returns>
        public ActionResult Selectll(BsgridPage bsgridPage, int? MessageID, string TissueCode, string ChineseName, string ChineseAbbreviation, string EnglishCode, string EnglishName, string Site, string Relation, string Describe, int? A)//查询组织
        {
            var ccc = (from tbMessage in myModels.SYS_Message
                       select new cctv
                       {
                           MessageID = tbMessage.MessageID,
                           TissueCode = tbMessage.TissueCode,//组织代码
                           ChineseName = tbMessage.ChineseName,//中文名称
                           ChineseAbbreviation = tbMessage.ChineseAbbreviation,//中文简称
                           Site = tbMessage.Site,//地址
                           EnglishCode = tbMessage.EnglishCode,//英文代码
                           EnglishName = tbMessage.EnglishName,//英文名称
                           Relation = tbMessage.Relation,//联系人
                           Describe = tbMessage.Describe,//描述
                       }).ToList();

            if (!string.IsNullOrEmpty(TissueCode))
            {
                ccc = ccc.Where(m => m.TissueCode.Contains(TissueCode)).ToList();//组织代码
            }
            if (!string.IsNullOrEmpty(ChineseName))
            {
                ccc = ccc.Where(m => m.ChineseName.Contains(ChineseName)).ToList();//中文名称
            }
            if (!string.IsNullOrEmpty(ChineseAbbreviation))
            {
                ccc = ccc.Where(m => m.ChineseAbbreviation.Contains(ChineseAbbreviation)).ToList();//中文简称
            }
            if (!string.IsNullOrEmpty(EnglishCode))
            {
                ccc = ccc.Where(m => m.EnglishCode.Contains(EnglishCode)).ToList();//英文代码
            }
            if (!string.IsNullOrEmpty(EnglishName))
            {
                ccc = ccc.Where(m => m.EnglishName.Contains(EnglishName)).ToList();//英文名称
            }
            if (!string.IsNullOrEmpty(Site))
            {
                ccc = ccc.Where(m => m.Site.Contains(Site)).ToList();//地址
            }
            if (!string.IsNullOrEmpty(Relation))
            {
                ccc = ccc.Where(m => m.Relation.Contains(Relation)).ToList();//联系人
            }
            if (!string.IsNullOrEmpty(Describe))
            {
                ccc = ccc.Where(m => m.Describe.Contains(Describe)).ToList();//描述
            }
            if (A == 1)
            {
                return Json(ccc, JsonRequestBehavior.AllowGet);
            }
            int pageCount = ccc.Count();
            List<cctv> wcg = ccc.Skip(bsgridPage.GetStartIndex()).Take(bsgridPage.pageSize).ToList();
            Bsgrid<cctv> bsgrid = new Bsgrid<cctv>()
            {
                totalRows = pageCount,
                success = true,
                curPage = bsgridPage.curPage,
                data = wcg
            };
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 新增
        /// </summary>
        /// <returns></returns>
        public ActionResult Inserll(string TissueCode, string ChineseName, string ChineseAbbreviation, string EnglishCode, string EnglishName, string Site, string Relation, string Describe)
        {
            try
            {
                //实例化一张表
                SYS_Message myMessage = new SYS_Message();
                //将页面写入的数据放到实例化的实体中
                myMessage.TissueCode = TissueCode.Trim();//组织代码
                myMessage.ChineseName = ChineseName.Trim();//中文名称
                myMessage.ChineseAbbreviation = ChineseAbbreviation.Trim();//中文简称
                myMessage.EnglishCode = EnglishCode.Trim();//英文代码
                myMessage.EnglishName = EnglishName.Trim();//英文名称
                myMessage.Site = Site.Trim();//地址
                myMessage.Relation = Relation.Trim();//联系人
                myMessage.Describe = Describe.Trim();//描述
                                                     //将实例化的实体插入数据库实体
                myModels.SYS_Message.Add(myMessage);
                //保存发生变化后的数据库
                myModels.SaveChanges();
                return Json("success", JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json("failed", JsonRequestBehavior.AllowGet);
            }
        }

        /// <summary>
        /// 修改
        /// </summary>
        /// <returns></returns>
        public ActionResult UpDate(short MessageID, string TissueCode, string ChineseName, string ChineseAbbreviation, string EnglishCode, string EnglishName, string Site, string Relation, string Describe)
        {
            string strMsg = "failed";
            try
            {
                Models.SYS_Message myMessage = myModels.SYS_Message.Where(p => p.MessageID == MessageID).Single();
                myMessage.TissueCode = TissueCode;//组织代码
                myMessage.ChineseName = ChineseName;//中文名称
                myMessage.ChineseAbbreviation = ChineseAbbreviation;//中文简称
                myMessage.EnglishCode = EnglishCode;//英文代码
                myMessage.EnglishName = EnglishName;//英文名称
                myMessage.Site = Site;//地址
                myMessage.Relation = Relation;//联系人
                myMessage.Describe = Describe;//描述       
                //获取和实质对象实体的状态 = EntityState的枚举值
                myModels.Entry(myMessage).State = EntityState.Modified;
                if (myModels.SaveChanges() > 0)
                {
                    strMsg = "success";
                    //      DataViewSettingCollection

                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);

        }

        /// <summary>
        /// 删除
        /// </summary>
        /// <param name="MessageID"></param>
        /// <returns></returns>
        public ActionResult Delete(short MessageID)
        {
            string strMsg = "fail";
            try
            {
                var Message = myModels.SYS_Message.Where(m => m.MessageID == MessageID).Single();
                myModels.SYS_Message.Remove(Message);
                if (myModels.SaveChanges() > 0)
                {
                    strMsg = "true";
                }

            }
            catch (Exception)
            {
                strMsg = "fail";

            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }



        /// <summary>
        /// 导出
        /// </summary>
        /// <returns></returns>
        public ActionResult FlowLeave(
           string TissueCode, string ChineseName, string ChineseAbbreviation, string EnglishCode, string EnglishName, string Site, string Relation, string Describe
            )
        {
            var ccc = (from tbMessage in myModels.SYS_Message
                       select new
                       {
                           MessageID = tbMessage.MessageID,
                           TissueCode = tbMessage.TissueCode,//组织代码
                           ChineseName = tbMessage.ChineseName,//中文名称
                           ChineseAbbreviation = tbMessage.ChineseAbbreviation,//中文简称
                           Site = tbMessage.Site,//地址
                           EnglishCode = tbMessage.EnglishCode,//英文代码
                           EnglishName = tbMessage.EnglishName,//英文名称
                           Relation = tbMessage.Relation,//联系人
                           Describe = tbMessage.Describe,//描述
                       }).ToList();
            if (!string.IsNullOrEmpty(TissueCode))
            {
                ccc = ccc.Where(m => m.TissueCode.Contains(TissueCode)).ToList();//组织代码
            }
            if (!string.IsNullOrEmpty(ChineseName))
            {
                ccc = ccc.Where(m => m.ChineseName.Contains(ChineseName)).ToList();//中文名称
            }
            if (!string.IsNullOrEmpty(ChineseAbbreviation))
            {
                ccc = ccc.Where(m => m.ChineseAbbreviation.Contains(ChineseAbbreviation)).ToList();//中文简称
            }
            if (!string.IsNullOrEmpty(EnglishCode))
            {
                ccc = ccc.Where(m => m.EnglishCode.Contains(EnglishCode)).ToList();//英文代码
            }
            if (!string.IsNullOrEmpty(EnglishName))
            {
                ccc = ccc.Where(m => m.EnglishName.Contains(EnglishName)).ToList();//英文名称
            }
            if (!string.IsNullOrEmpty(Site))
            {
                ccc = ccc.Where(m => m.Site.Contains(Site)).ToList();//地址
            }
            if (!string.IsNullOrEmpty(Relation))
            {
                ccc = ccc.Where(m => m.Relation.Contains(Relation)).ToList();//联系人
            }
            if (!string.IsNullOrEmpty(Describe))
            {
                ccc = ccc.Where(m => m.Describe.Contains(Describe)).ToList();//描述
            }
            //1、创建Excel工作簿
            NPOI.HSSF.UserModel.HSSFWorkbook Leave = new NPOI.HSSF.UserModel.HSSFWorkbook();
            //   NPOI.HSSF.UserModel.HSSFWorkbook exBook = new NPOI.HSSF.UserModel.HSSFWorkbook();

            //2、创建工作表
            NPOI.SS.UserModel.ISheet sheet = Leave.CreateSheet("东郭先生");
            //3、设计表头
            //3.1 创建第一行
            NPOI.SS.UserModel.IRow row = sheet.CreateRow(0);
            //3.2 设计字段
            row.CreateCell(0).SetCellValue("组织代码");
            row.CreateCell(1).SetCellValue("中文名称");
            row.CreateCell(2).SetCellValue("中文简称");
            row.CreateCell(3).SetCellValue("英文代码");
            row.CreateCell(4).SetCellValue("英文名称");
            row.CreateCell(5).SetCellValue("地址");
            row.CreateCell(6).SetCellValue("联系人");
            row.CreateCell(7).SetCellValue("描述");
            //4、循环写入数据
            for (var i = 0; i < ccc.Count(); i++)
            {
                NPOI.SS.UserModel.IRow rowl = sheet.CreateRow(i + 1);
                rowl.CreateCell(0).SetCellValue(ccc[i].TissueCode);
                rowl.CreateCell(1).SetCellValue(ccc[i].ChineseName);
                rowl.CreateCell(2).SetCellValue(ccc[i].ChineseAbbreviation);
                rowl.CreateCell(3).SetCellValue(ccc[i].EnglishCode);
                rowl.CreateCell(3).SetCellValue(ccc[i].EnglishName);
                rowl.CreateCell(4).SetCellValue(ccc[i].Site);
                rowl.CreateCell(5).SetCellValue(ccc[i].Relation);
                rowl.CreateCell(6).SetCellValue(ccc[i].Describe);
            }
            //5、将Excel表格转化为文件流输出
            MemoryStream DulaDula = new MemoryStream();
            Leave.Write(DulaDula);
            //输出之前调用Seek（偏移量，游标位置）方法：确定流开始的位置
            DulaDula.Seek(0, SeekOrigin.Begin);
            //6、文件名
            string fileName = "东郭先生" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            return File(DulaDula, "application/vnd.ms-excel", fileName);

        }
        /// <summary>
        /// 打印
        /// </summary>
        /// <returns></returns>
        public ActionResult Stamp(string TissueCode, string ChineseName, string ChineseAbbreviation, string EnglishCode, string EnglishName, string Site, string Relation, string Describe)
        {
            #region 
            var ccc = (from tbMessage in myModels.SYS_Message
                       select new
                       {
                           MessageID = tbMessage.MessageID,
                           TissueCode = tbMessage.TissueCode,//组织代码
                           ChineseName = tbMessage.ChineseName,//中文名称
                           ChineseAbbreviation = tbMessage.ChineseAbbreviation,//中文简称
                           Site = tbMessage.Site,//地址
                           EnglishCode = tbMessage.EnglishCode,//英文代码
                           EnglishName = tbMessage.EnglishName,//英文名称
                           Relation = tbMessage.Relation,//联系人
                           Describe = tbMessage.Describe,//描述
                       }).ToList();
            if (!string.IsNullOrEmpty(TissueCode))
            {
                ccc = ccc.Where(m => m.TissueCode.Contains(TissueCode)).ToList();//组织代码
            }
            if (!string.IsNullOrEmpty(ChineseName))
            {
                ccc = ccc.Where(m => m.ChineseName.Contains(ChineseName)).ToList();//中文名称
            }
            if (!string.IsNullOrEmpty(ChineseAbbreviation))
            {
                ccc = ccc.Where(m => m.ChineseAbbreviation.Contains(ChineseAbbreviation)).ToList();//中文简称
            }
            if (!string.IsNullOrEmpty(EnglishCode))
            {
                ccc = ccc.Where(m => m.EnglishCode.Contains(EnglishCode)).ToList();//英文代码
            }
            if (!string.IsNullOrEmpty(EnglishName))
            {
                ccc = ccc.Where(m => m.EnglishName.Contains(EnglishName)).ToList();//英文名称
            }
            if (!string.IsNullOrEmpty(Site))
            {
                ccc = ccc.Where(m => m.Site.Contains(Site)).ToList();//地址
            }
            if (!string.IsNullOrEmpty(Relation))
            {
                ccc = ccc.Where(m => m.Relation.Contains(Relation)).ToList();//联系人
            }
            if (!string.IsNullOrEmpty(Describe))
            {
                ccc = ccc.Where(m => m.Describe.Contains(Describe)).ToList();//描述
            }
            #endregion
            //将查询出来的结果转化为DataTable的格式
            DataTable dtResult = LINQToDataTable(ccc);
            //1、实例化数据集aa
            pp.ccc dbReport = new pp.ccc();
            //2、将dtResult放入数据集中名为"Message"的表格中  与数据集的表相同
            dbReport.Tables["Message"].Merge(dtResult);
            //3、实例化数据报表
            pp.Message rp = new pp.Message();
            //4、获取报表的物理文件路径
            string strRpPath = System.Web.HttpContext.Current.Server.MapPath("~/") + "Areas\\PPAP\\pp\\BB.rpt";
            //5、将报表加载到报表模板中
            rp.Load(strRpPath);
            //6、设置报表的数据源
            rp.SetDataSource(dbReport);
            //7、把报表转化为文件流输出
            Stream dbStream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
            return File(dbStream, "application/pdf");
        }
        /// <summary>
        /// 接口
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="varlist"></param>
        /// <returns></returns>
        public DataTable LINQToDataTable<T>(IEnumerable<T> varlist)
        {//定义要返回的DataTable对象
            DataTable dtReturn = new DataTable();
            //DataTable ggg = new DataTable();
            //保存列集合的属性信息数组
            PropertyInfo[] oProps = null;
            if (varlist == null)
                return dtReturn;//安全性检查
            //循环遍历集合，使用反射获取类型的属性信息
            foreach (T rec in varlist)
            {
                //使用反射获取T类型的属性信息，返回一个PropertyInfo类型的集合
                #region 
                if (oProps == null)
                {
                    oProps = ((Type)rec.GetType()).GetProperties();
                    //循环PropertyInfo数组
                    foreach (PropertyInfo pi in oProps)
                    {
                        //得到属性的类
                        Type colType = pi.PropertyType;
                        //如果属性为泛型类型

                        if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition() == typeof(Nullable<>)))
                        {
                            //获取泛型类型的参数
                            colType = colType.GetGenericArguments()[0];
                        }
                        //将类型的属性名称与属性类型作为DataTable的列数据
                        dtReturn.Columns.Add(pi.Name, colType);
                    }
                }
                #endregion
                //新建一个用于添加到DataTable中的DataRow对象
                DataRow dr = dtReturn.NewRow();
                //循环遍历属性集合
                foreach (PropertyInfo pi in oProps)
                {
                    //为DataRow中的指定列赋值
                    dr[pi.Name] = pi.GetValue(rec, null) == null ? DBNull.Value : pi.GetValue(rec, null);
                }
                //将具有结果值的DataRow添加到DataTable集合中
                dtReturn.Rows.Add(dr);
            }
            return dtReturn;//返回DataTable对象
        }

        #endregion

    }
}
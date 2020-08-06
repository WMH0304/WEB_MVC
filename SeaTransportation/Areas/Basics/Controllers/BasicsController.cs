using Hotel.VO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using SeaTransportation.Models;
using SeaTransportation.Vo;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Web.Mvc;

namespace SeaTransportation.Areas.Basics.Controllers
{
    public class BasicsController : Controller
    {
        Models.SeaTransportationEntities myModels = new Models.SeaTransportationEntities();
        // GET: Basics/Basics

        #region 页面
        public ActionResult Ship()//船舶资料
        {
            //重定向代码
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");
            }
            return View();
        }
        public ActionResult Billing()//计费门点
        {
            //重定向代码
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");
            }
            return View();
        }
        public ActionResult Chauffeur()//司机信息
        {
            //重定向代码
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");
            }
            return View();
        }
        public ActionResult DriverBook()//司机本
        {
            //重定向代码
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");
            }
            return View();
        }
        public ActionResult VehicleInformation()//车辆信息
        {
            //重定向代码
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");
            }
            return View();
        }
        public ActionResult Bracket()//托架资料
        {
            //重定向代码
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");
            }
            return View();
        }
        public ActionResult Mention()//提还柜地
        {
            //重定向代码
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");
            }
            return View();
        }
        public ActionResult Customs()//关区管理
        {
            //重定向代码
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");
            }
            return View();
        }
        public ActionResult HaulWayy()//运输路线
        {
            //重定向代码
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");
            }
            return View();
        }     
        public ActionResult Expense()//费用项目
        {
            //重定向代码
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");
            }
            return View();
        }
        public ActionResult Parities()//系统汇率
        {
            //重定向代码
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");
            }
            return View();
        }
        public ActionResult Port()//港口资料
        {
            //重定向代码
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");
            }
            return View();
        }
        #endregion

        #region  下拉框↓
        /// <summary>
        /// 查询船名
        /// </summary>
        /// <returns></returns>
        public ActionResult selectshi()
        {
            var list = from tbClientType in myModels.SYS_ClientType
                       join tbClient in myModels.SYS_Client on tbClientType.ClientID equals tbClient.ClientID
                       where tbClientType.ClientType.Trim() == "船公司"
                       select new
                       {
                           id = tbClientType.ClientTypeID,
                           name = tbClient.ChineseName
                       };
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 查询公司
        /// </summary>
        /// <returns></returns>
        public ActionResult selecttmMessage()
        {
            var list = myModels.SYS_Message.Select(m => new { id = m.MessageID, name = m.ChineseName });
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 查询司机
        /// </summary>
        /// <returns></returns>
        public ActionResult selectchauffer()
        {
            var list = myModels.SYS_Chauffeur.Select(m => new { id = m.ChauffeurID, name = m.ChauffeurName }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 查询托架
        /// </summary>
        /// <returns></returns>
        public ActionResult selectbracket()
        {
            var list = myModels.SYS_Bracket.Select(m => new { id = m.BracketID, name = m.BracketCode }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 查询门点
        /// </summary>
        /// <returns></returns>
        public ActionResult selectGatedotxia()
        {
            var list = myModels.SYS_Gatedot.Select(m => new { id = m.GatedotID, name = m.GatedotName }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 查询提还
        /// </summary>
        /// <returns></returns>
        public ActionResult selectmentionxia()
        {
            var list = myModels.SYS_Mention.Select(m => new { id = m.MentionID, name = m.Abbreviation }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 查询报关
        /// </summary>
        /// <returns></returns>
        public ActionResult selectcustomsxia()
        {
            var list = myModels.SYS_Customs.Select(m => new { id = m.CustomsID, name = m.CustomsName }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 查询员工
        /// </summary>
        /// <returns></returns>
        public ActionResult staff()
        {
            var list = myModels.SYS_Staff.Select(m => new { id = m.StaffID, name = m.StaffName });
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 查询路线
        /// </summary>
        /// <returns></returns>
        public ActionResult haulway()
        {
            var list = myModels.SYS_HaulWay.Select(m => new { id = m.HaulWayID,name = m.HaulWayDescription });
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 查询费用
        /// </summary>
        /// <returns></returns>
        public ActionResult expensee()
        {
            var list = myModels.SYS_Expense.Select(m => new { id = m.ExpenseID, name = m.ExpenseName });
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 查询报价
        /// </summary>
        /// <returns></returns>
        public ActionResult offer()
        {
            var list = from tbOffer in myModels.SYS_Offer
                       where tbOffer.OfferType.Trim() == "司机产值"
                       select new
                       {
                           id = tbOffer.OfferID,
                           name = tbOffer.OfferType
                       };
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region  船舶资料
        /// <summary>
        /// 查询船舶
        /// </summary>
        /// <returns></returns>
        public ActionResult selectShip(BsgridPage bsgridPage, string ShipCode, string Abbreviation, string ChineseName, string EnglishName,string Operator, string ShipEast, string ShipType, string CaptainName, string ShipMobile, string ChineseNamel,int? ClientTypeID)
        {
            var list = (from tbShip in myModels.SYS_Ship
                        join tbClientType in myModels.SYS_ClientType on tbShip.ClientTypeID equals tbClientType.ClientTypeID
                        join tbClient in myModels.SYS_Client on tbClientType.ClientID equals tbClient.ClientID
                        select new Ship
                        {
                            ClientID = tbClient.ClientID,
                            ChineseNamel = tbClient.ChineseName,
                            ShipID = tbShip.ShipID,
                            ClientTypeID = tbShip.ClientTypeID,
                            ChineseName = tbShip.ChineseName,
                            ClientType = tbClientType.ClientType,
                            ShipCode = tbShip.ShipCode,
                            EnglishName = tbShip.EnglishName,
                            Abbreviation = tbShip.Abbreviation,
                            ShipEast = tbShip.ShipEast,
                            ShipType = tbShip.ShipType.Trim(),
                            CaptainName = tbShip.CaptainName,
                            ShipMobile = tbShip.ShipMobile,
                            WhetherStart = tbShip.WhetherStart,
                            Operator = tbShip.Operator
                        }).ToList();
            if (!string.IsNullOrEmpty(ShipCode))
            {
                list = list.Where(s => s.ShipCode.Contains(ShipCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(Abbreviation))
            {
                list = list.Where(s => s.Abbreviation.Contains(Abbreviation.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ChineseName))
            {
                list = list.Where(s => s.ChineseName.Contains(ChineseName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(EnglishName))
            {
                list = list.Where(s => s.EnglishName.Contains(EnglishName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ChineseNamel))
            {
                list = list.Where(s => s.ChineseNamel.Contains(ChineseNamel.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(Operator))
            {
                list = list.Where(s => s.Operator.Contains(Operator.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ShipEast))
            {
                list = list.Where(s => s.ShipEast.Contains(ShipEast.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ShipType))
            {
                list = list.Where(s => s.ShipType.Contains(ShipType.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(CaptainName))
            {
                list = list.Where(s => s.CaptainName.Contains(CaptainName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ShipMobile))
            {
                list = list.Where(s => s.ShipMobile.Contains(ShipMobile.Trim())).ToList();
            }
            if (ClientTypeID>0)
            {
                list = list.Where(s => s.ClientTypeID== ClientTypeID).ToList();
            }
            int totalRow = list.Count();
            List<Ship> notices = list.OrderByDescending(p => p.ShipID).
              Skip(bsgridPage.GetStartIndex()).
              Take(bsgridPage.pageSize).
              ToList();
            Bsgrid<Ship> bsgrid = new Bsgrid<Ship>();
            bsgrid.success = true;
            bsgrid.totalRows = totalRow;
            bsgrid.curPage = bsgridPage.curPage;
            bsgrid.data = notices;
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 新增船舶
        /// </summary>
        /// <returns></returns>
        public ActionResult InsertShip(SYS_Ship tbShip)
        {
            string strMsg = "fail";
            try
            {
                int oldship = (from tbship in myModels.SYS_Ship
                               where tbship.ShipCode == tbShip.ShipCode
                               select tbship).Count();
                if (oldship == 0)
                {
                    int ship = (from tbship in myModels.SYS_Ship
                                where tbship.ChineseName == tbShip.ChineseName
                                select tbship).Count();
                    if (ship == 0)
                    {
                        myModels.SYS_Ship.Add(tbShip);
                        if (myModels.SaveChanges() > 0)
                        {
                            strMsg = "数据新增成功！";
                        }
                    }
                    else
                    {
                        strMsg = "已有中文名,无法新增！";
                    }
                }
                else
                {
                    strMsg = "已有船舶代码,无法新增！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 修改船舶
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateShip(SYS_Ship pwShip)
        {
            string strMsg = "fail";
            try
            {
                if (myModels.SYS_Ship.Where(m => m.ShipCode == pwShip.ShipCode && m.ShipID != pwShip.ShipID).Count() == 0)
                {
                    if (myModels.SYS_Ship.Where(m => m.ChineseName == pwShip.ChineseName && m.ShipID != pwShip.ShipID).Count() == 0)
                    {
                        SYS_Ship dbShip = (from tbShip in myModels.SYS_Ship
                                           where tbShip.ShipID == pwShip.ShipID
                                           select tbShip).Single();
                        dbShip.ClientTypeID = pwShip.ClientTypeID;
                        dbShip.Abbreviation = pwShip.Abbreviation;
                        dbShip.CaptainName = pwShip.CaptainName;
                        dbShip.ChineseName = pwShip.ChineseName;
                        dbShip.EnglishName = pwShip.EnglishName;
                        dbShip.Operator = pwShip.Operator;
                        dbShip.ShipCode = pwShip.ShipCode;
                        dbShip.ShipEast = pwShip.ShipEast;
                        dbShip.ShipMobile = pwShip.ShipMobile;
                        dbShip.ShipType = pwShip.ShipType;
                        dbShip.WhetherStart = pwShip.WhetherStart;
                        myModels.Entry(dbShip).State = System.Data.Entity.EntityState.Modified;
                        if (myModels.SaveChanges() > 0)
                        {
                            strMsg = "修改成功！";
                        }
                    }
                    else
                    {
                        strMsg = "已有中文名，无法修改！";
                    }
                }
                else
                {
                    strMsg = "已有船舶代码，无法修改！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 回填船舶
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectshipById(int ShipID)
        {
            try
            {
                var list = (from tbShip in myModels.SYS_Ship
                            join tbClientType in myModels.SYS_ClientType on tbShip.ClientTypeID equals tbClientType.ClientTypeID
                            join tbClient in myModels.SYS_Client on tbClientType.ClientID equals tbClient.ClientID
                            where tbShip.ShipID == ShipID
                            select new Ship
                            {
                                ClientID = tbClient.ClientID,
                                ChineseNamel = tbClient.ChineseName.Trim(),
                                ShipID = tbShip.ShipID,
                                ClientTypeID = tbShip.ClientTypeID,
                                ChineseName = tbShip.ChineseName.Trim(),
                                ClientType = tbClientType.ClientType,
                                ShipCode = tbShip.ShipCode,
                                EnglishName = tbShip.EnglishName,
                                Abbreviation = tbShip.Abbreviation,
                                ShipEast = tbShip.ShipEast,
                                ShipType = tbShip.ShipType.Trim(),
                                CaptainName = tbShip.CaptainName,
                                ShipMobile = tbShip.ShipMobile,
                                WhetherStart = tbShip.WhetherStart,
                                Operator = tbShip.Operator
                            }).Single();
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return null;
            }
        }
        /// <summary>
        /// 删除船舶
        /// </summary>
        /// <returns></returns>
        public ActionResult DelectShip(int ShipID)
        {
            string strMsg = "fail";
            try
            {
                var listShip = myModels.SYS_Ship.Where(m => m.ShipID == ShipID).Single();
                myModels.SYS_Ship.Remove(listShip);
                if (myModels.SaveChanges() > 0)
                {
                    strMsg = "true";
                }
            }
            catch (Exception e)
            {
                strMsg = "fail";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 导出船舶
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportToExcel(string ShipCode, string ChineseName, string EnglishName, string Abbreviation, string ShipEast, string ShipType, string CaptainName, string ShipMobile, string Operator, string ChineseNamel)
        {
            var list = (from tbShip in myModels.SYS_Ship
                        join tbClientType in myModels.SYS_ClientType on tbShip.ClientTypeID equals tbClientType.ClientTypeID
                        join tbClient in myModels.SYS_Client on tbClientType.ClientID equals tbClient.ClientID
                        select new Ship
                        {
                            ClientID = tbClient.ClientID,
                            ChineseNamel = tbClient.ChineseName,
                            ShipID = tbShip.ShipID,
                            ClientTypeID = tbClientType.ClientTypeID,
                            ChineseName = tbShip.ChineseName,
                            ClientType = tbClientType.ClientType,
                            ShipCode = tbShip.ShipCode,
                            EnglishName = tbShip.EnglishName,
                            Abbreviation = tbShip.Abbreviation,
                            ShipEast = tbShip.ShipEast,
                            ShipType = tbShip.ShipType,
                            CaptainName = tbShip.CaptainName,
                            ShipMobile = tbShip.ShipMobile,
                            WhetherStart = tbShip.WhetherStart,
                            Operator = tbShip.Operator
                        }).ToList();
            if (!string.IsNullOrEmpty(ChineseNamel))
            {
                list = list.Where(s => s.ChineseNamel.Contains(ChineseNamel.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ShipCode))
            {
                list = list.Where(s => s.ShipCode.Contains(ShipCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ChineseName))
            {
                list = list.Where(s => s.ChineseName.Contains(ChineseName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(EnglishName))
            {
                list = list.Where(s => s.EnglishName.Contains(EnglishName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(Abbreviation))
            {
                list = list.Where(s => s.Abbreviation.Contains(Abbreviation.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ShipEast))
            {
                list = list.Where(s => s.ShipEast.Contains(ShipEast.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ShipType))
            {
                list = list.Where(s => s.ShipType.Contains(ShipType.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(CaptainName))
            {
                list = list.Where(s => s.CaptainName.Contains(CaptainName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ShipMobile))
            {
                list = list.Where(s => s.ShipMobile.Contains(ShipMobile.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(Operator))
            {
                list = list.Where(s => s.Operator.Contains(Operator.Trim())).ToList();
            }
            //查询数据
            List<Ship> listExaminee = list.ToList();
            //二：代码创建一个Excel表格（这里称为工作簿）
            //创建Excel文件的对象 工作簿(调用NPOI文件)
            HSSFWorkbook excelBook = new HSSFWorkbook();
            //创建Excel工作表 Sheet
            ISheet sheet1 = excelBook.CreateSheet("船舶资料");
            //给Sheet添加第一行的头部标题
            IRow row1 = sheet1.CreateRow(0);
            //给标题的每一个单元格赋值
            row1.CreateCell(0).SetCellValue("船舶代码");
            row1.CreateCell(1).SetCellValue("中文名");
            row1.CreateCell(2).SetCellValue("英文名");
            row1.CreateCell(3).SetCellValue("简称");
            row1.CreateCell(4).SetCellValue("船公司");
            row1.CreateCell(5).SetCellValue("船东");
            row1.CreateCell(6).SetCellValue("船舶类型");
            row1.CreateCell(7).SetCellValue("船长姓名");
            row1.CreateCell(8).SetCellValue("联系电话");
            row1.CreateCell(9).SetCellValue("是否启用");
            row1.CreateCell(10).SetCellValue("经营者");
            //添加数据行：将表格数据逐步写入sheet1各个行中（也就是给每一个单元格赋值）
            for (int i = 0; i < listExaminee.Count; i++)
            {
                //创建行
                IRow rowTemp = sheet1.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listExaminee[i].ShipCode);
                rowTemp.CreateCell(1).SetCellValue(listExaminee[i].ChineseName);
                rowTemp.CreateCell(2).SetCellValue(listExaminee[i].EnglishName);
                rowTemp.CreateCell(3).SetCellValue(listExaminee[i].Abbreviation);
                rowTemp.CreateCell(4).SetCellValue(listExaminee[i].ChineseNamel);
                rowTemp.CreateCell(5).SetCellValue(listExaminee[i].ShipEast);
                rowTemp.CreateCell(6).SetCellValue(listExaminee[i].ShipType);
                rowTemp.CreateCell(7).SetCellValue(listExaminee[i].CaptainName);
                rowTemp.CreateCell(8).SetCellValue(listExaminee[i].ShipMobile);
                rowTemp.CreateCell(9).SetCellValue(listExaminee[i].WhetherStart);
                rowTemp.CreateCell(10).SetCellValue(listExaminee[i].Operator);
            }
            //输出的文件名称
            string fileName = "船舶资料" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
            //把Excel转为流，输出
            //创建文件流
            System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
            //将工作薄写入文件流
            excelBook.Write(bookStream);
            //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
            bookStream.Seek(0, System.IO.SeekOrigin.Begin);
            //Stream对象,文件类型,文件名称
            return File(bookStream, "application/vnd.ms-excel", fileName);
        }
        /// <summary>
        /// 船舶报表
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievement(string ShipCode, string ChineseName, string EnglishName, string Abbreviation, string ShipEast, string ShipType, string CaptainName, string ShipMobile, string Operator, string ChineseNamel)
        {
            var list = (from tbShip in myModels.SYS_Ship
                        join tbClientType in myModels.SYS_ClientType on tbShip.ClientTypeID equals tbClientType.ClientTypeID
                        join tbClient in myModels.SYS_Client on tbClientType.ClientID equals tbClient.ClientID
                        select new Ship
                        {
                            ClientID = tbClient.ClientID,
                            ChineseNamel = tbClient.ChineseName,
                            ShipID = tbShip.ShipID,
                            ClientTypeID = tbClientType.ClientTypeID,
                            ChineseName = tbShip.ChineseName,
                            ClientType = tbClientType.ClientType,
                            ShipCode = tbShip.ShipCode,
                            EnglishName = tbShip.EnglishName,
                            Abbreviation = tbShip.Abbreviation,
                            ShipEast = tbShip.ShipEast,
                            ShipType = tbShip.ShipType,
                            CaptainName = tbShip.CaptainName,
                            ShipMobile = tbShip.ShipMobile,
                            WhetherStarta = "",
                            WhetherStartl = tbShip.WhetherStart,
                            Operator = tbShip.Operator
                        }).ToList();
            if (!string.IsNullOrEmpty(ChineseNamel))
            {
                list = list.Where(s => s.ChineseNamel.Contains(ChineseNamel.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ShipCode))
            {
                list = list.Where(s => s.ShipCode.Contains(ShipCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ChineseName))
            {
                list = list.Where(s => s.ChineseName.Contains(ChineseName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(EnglishName))
            {
                list = list.Where(s => s.EnglishName.Contains(EnglishName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(Abbreviation))
            {
                list = list.Where(s => s.Abbreviation.Contains(Abbreviation.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ShipEast))
            {
                list = list.Where(s => s.ShipEast.Contains(ShipEast.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ShipType))
            {
                list = list.Where(s => s.ShipType.Contains(ShipType.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(CaptainName))
            {
                list = list.Where(s => s.CaptainName.Contains(CaptainName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ShipMobile))
            {
                list = list.Where(s => s.ShipMobile.Contains(ShipMobile.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(Operator))
            {
                list = list.Where(s => s.Operator.Contains(Operator.Trim())).ToList();
            }
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i].WhetherStartl == true)
                {
                    list[i].WhetherStarta = "√";
                }
                else if (list[i].WhetherStartl == false)
                {
                    list[i].WhetherStarta = "×";
                }
            }
            
            //把linq类型的数据listResult转化为DataTable类型数据
            DataTable dt = LINQToDataTable(list);
            //第一步：实例化数据集`
            PrintReport.ReportDB dbReport = new PrintReport.ReportDB();
            //第二步：将dt的数据放入数据集的数据表中
            dbReport.Tables["Ship"].Merge(dt);
            //第三步：实例化报表模板
            PrintReport.ship rp = new PrintReport.ship();
            //第四步：获取报表物理文件地址     
            string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                + "Areas\\Basics\\PrintReport\\ship.rpt";
            //第五步：把报表文件加载到ReportDocument
            rp.Load(strRptPath);
            //第六步：设置报表数据源
            rp.SetDataSource(dbReport);
            //第七步：把ReportDocument转化为文件流
            Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
            return File(stream, "application/pdf");
        }
        // //将IEnumerable<T>类型的集合转换为DataTable类型
        public DataTable LINQToDataTable<T>(IEnumerable<T> varlist)
        {
            //定义要返回的DataTable对象
            DataTable dtReturn = new DataTable();
            //保存列集合的属性信息数组
            PropertyInfo[] oProps = null;
            if (varlist == null) return dtReturn;//安全性检查
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
                        //得到属性的类型
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
        #region 计费门点
        /// <summary>
        /// 查询城市
        /// </summary>
        /// <returns></returns>
        public ActionResult selectcityxia(string Province)
        {
            var list = myModels.BS_City.Where(p => p.Province == Province).Select(m => new { id = m.CityID, name = m.CityName });
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        public ActionResult selectcityxiaxia(string Province)
        {
            if (!string.IsNullOrEmpty(Province))
            {
                var list = myModels.BS_City.Where(p => p.Province == Province).Select(m => new { id = m.CityID, name = m.CityName });
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            else
            {
                var list = myModels.BS_City.Select(m => new { id = m.CityID, name = m.CityName });
                return Json(list, JsonRequestBehavior.AllowGet);
            }
        }
        /// <summary>
        /// 查询城市
        /// </summary>
        /// <returns></returns>
        public ActionResult selectbilling(BsgridPage bsgridPage, string CityName, string Province)
        {
            var list = (from tbCity in myModels.BS_City
                        select new City
                        {
                            CityID = tbCity.CityID,
                            Province = tbCity.Province,
                            CityCode = tbCity.CityCode,
                            CityName = tbCity.CityName,
                            Describe = tbCity.Describe,
                            EnglishName = tbCity.EnglishName
                        }).ToList();
            if (!string.IsNullOrEmpty(CityName))
            {
                list = list.Where(m => m.CityName.Contains(CityName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(Province))
            {
                list = list.Where(m => m.Province.Contains(Province.Trim())).ToList();
            }
            int totalRow = list.Count();
            List<City> notices = list.OrderByDescending(p => p.CityID).
              Skip(bsgridPage.GetStartIndex()).
              Take(bsgridPage.pageSize).
              ToList();
            Bsgrid<City> bsgrid = new Bsgrid<City>();
            bsgrid.success = true;
            bsgrid.totalRows = totalRow;
            bsgrid.curPage = bsgridPage.curPage;
            bsgrid.data = notices;
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 查询门点
        /// </summary>
        /// <returns></returns>
        public ActionResult selectGatedot(BsgridPage bsgridPage, int CityID)
        {
            var list = (from tbGatedot in myModels.SYS_Gatedot
                        where tbGatedot.CityID == CityID
                        select new Gatedot
                        {
                            GatedotID = tbGatedot.GatedotID,
                            CityID = tbGatedot.CityID,
                            GatedotCode = tbGatedot.GatedotCode,
                            GatedotName = tbGatedot.GatedotName,
                            Describe = tbGatedot.Describe,
                            WhetherValid = tbGatedot.WhetherValid,
                            WhetherSuperiorGatedot = tbGatedot.WhetherSuperiorGatedot
                        }).ToList();
            int totalRow = list.Count();
            List<Gatedot> notices = list.OrderByDescending(p => p.GatedotID).
              Skip(bsgridPage.GetStartIndex()).
              Take(bsgridPage.pageSize).
              ToList();
            Bsgrid<Gatedot> bsgrid = new Bsgrid<Gatedot>();
            bsgrid.success = true;
            bsgrid.totalRows = totalRow;
            bsgrid.curPage = bsgridPage.curPage;
            bsgrid.data = notices;
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 新增门点
        /// </summary>
        /// <returns></returns>
        public ActionResult InsertGatedot(SYS_Gatedot tbgatedot)
        {
            string strMsg = "fail";
            try
            {
                int oldGatedot = (from tbGatedot in myModels.SYS_Gatedot
                                  where tbGatedot.GatedotCode == tbgatedot.GatedotCode
                                  select tbGatedot).Count();
                if (oldGatedot == 0)
                {
                    int Gatedot = (from tbGatedot in myModels.SYS_Gatedot
                                   where tbGatedot.GatedotName == tbgatedot.GatedotName
                                   select tbGatedot).Count();
                    if (Gatedot == 0)
                    {
                        myModels.SYS_Gatedot.Add(tbgatedot);
                        if (myModels.SaveChanges() > 0)
                        {
                            strMsg = "新增成功！";
                        }
                    }
                    else
                    {
                        strMsg = "已有门点名称！";
                    }
                }
                else
                {
                    strMsg = "已有门点编码！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 修改门点
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateGatedot(SYS_Gatedot tbgatedot)
        {
            string strMsg = "fail";
            try
            {
                SYS_Gatedot dbGatedot = (from tbGatedot in myModels.SYS_Gatedot
                                         where tbGatedot.GatedotID == tbgatedot.GatedotID
                                         select tbGatedot).Single();
                dbGatedot.CityID = tbgatedot.CityID;
                dbGatedot.Describe = tbgatedot.Describe;
                dbGatedot.GatedotCode = tbgatedot.GatedotCode;
                dbGatedot.GatedotName = tbgatedot.GatedotName;
                dbGatedot.WhetherValid = tbgatedot.WhetherValid;
                dbGatedot.WhetherSuperiorGatedot = tbgatedot.WhetherSuperiorGatedot;
                myModels.Entry(dbGatedot).State = System.Data.Entity.EntityState.Modified;
                if (myModels.SaveChanges() > 0)
                {
                    strMsg = "修改成功！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 回填门点
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectGatedotById(int GatedotID)
        {
            try
            {
                var list = (from tbGatedot in myModels.SYS_Gatedot
                            join tb in myModels.BS_City on tbGatedot.CityID equals tb.CityID
                            where tbGatedot.GatedotID == GatedotID
                            select new Gatedot
                            {
                                CityID = tbGatedot.CityID,
                                Province = tb.Province.Trim(),
                                GatedotID = tbGatedot.GatedotID,
                                Describe = tbGatedot.Describe,
                                GatedotCode = tbGatedot.GatedotCode,
                                GatedotName = tbGatedot.GatedotName,
                                WhetherSuperiorGatedot = tbGatedot.WhetherSuperiorGatedot,
                                WhetherValid = tbGatedot.WhetherValid
                            }).Single();
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return null;
            }
        }
        /// <summary>
        /// 删除门点
        /// </summary>
        /// <returns></returns>
        public ActionResult DelectGatedot(int GatedotID)
        {
            string strMsg = "fail";
            try
            {
                var listGatedot = myModels.SYS_Gatedot.Where(m => m.GatedotID == GatedotID).Single();
                myModels.SYS_Gatedot.Remove(listGatedot);
                if (myModels.SaveChanges() > 0)
                {
                    strMsg = "true";
                }
            }
            catch (Exception e)
            {
                strMsg = "fail";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 导出门点
        /// </summary>
        /// <returns></returns>
        public ActionResult GatedotExportToExcel()
        {
            var list = (from tbGatedot in myModels.SYS_Gatedot
                        join tbCity in myModels.BS_City on tbGatedot.CityID equals tbCity.CityID
                        select new Gatedot
                        {
                            CityID = tbGatedot.CityID,
                            CityName = tbCity.CityName,
                            GatedotID = tbGatedot.GatedotID,
                            Describe = tbGatedot.Describe,
                            GatedotCode = tbGatedot.GatedotCode,
                            GatedotName = tbGatedot.GatedotName,
                            WhetherSuperiorGatedota="",
                            WhetherSuperiorGatedotb = tbGatedot.WhetherSuperiorGatedot,
                            WhetherValida="",
                            WhetherValidb = tbGatedot.WhetherValid,
                        }).ToList();
            for (int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherSuperiorGatedotb == true)
                {
                    list[i].WhetherSuperiorGatedota = "有上级门点";
                }
                else if (list[i].WhetherSuperiorGatedotb == false)
                {
                    list[i].WhetherSuperiorGatedota = "没有上级门点";
                }
            }
            for (int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherValidb == true)
                {
                    list[i].WhetherValida = "有效";
                }
                else if (list[i].WhetherValidb == false)
                {
                    list[i].WhetherValida = "无效";
                }
            }
            //查询数据
            List<Gatedot> listExaminee = list.ToList();
            //二：代码创建一个Excel表格（这里称为工作簿）
            //创建Excel文件的对象 工作簿(调用NPOI文件)
            HSSFWorkbook excelBook = new HSSFWorkbook();
            //创建Excel工作表 Sheet
            ISheet sheet1 = excelBook.CreateSheet("门点资料");
            //给Sheet添加第一行的头部标题
            IRow row1 = sheet1.CreateRow(0);
            //给标题的每一个单元格赋值
            row1.CreateCell(0).SetCellValue("城市名称");
            row1.CreateCell(1).SetCellValue("门点编码");
            row1.CreateCell(2).SetCellValue("门点名称");
            row1.CreateCell(3).SetCellValue("描述");
            row1.CreateCell(4).SetCellValue("是否有上级门点");
            row1.CreateCell(5).SetCellValue("是否有效");
            //添加数据行：将表格数据逐步写入sheet1各个行中（也就是给每一个单元格赋值）
            for (int i = 0; i < listExaminee.Count; i++)
            {
                IRow rowTemp = sheet1.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listExaminee[i].CityName.Trim());
                rowTemp.CreateCell(1).SetCellValue(listExaminee[i].GatedotCode.Trim());
                rowTemp.CreateCell(2).SetCellValue(listExaminee[i].GatedotName.Trim());
                rowTemp.CreateCell(3).SetCellValue(listExaminee[i].Describe);
                rowTemp.CreateCell(4).SetCellValue(listExaminee[i].WhetherSuperiorGatedota);
                rowTemp.CreateCell(5).SetCellValue(listExaminee[i].WhetherValida);
            }
            //输出的文件名称
            string fileName = "门点资料" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
            //把Excel转为流，输出
            //创建文件流
            System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
            //将工作薄写入文件流
            excelBook.Write(bookStream);
            //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
            bookStream.Seek(0, System.IO.SeekOrigin.Begin);
            //Stream对象,文件类型,文件名称
            return File(bookStream, "application/vnd.ms-excel", fileName);
        }
        /// <summary>
        /// 门点报表
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievementGatedot()
        {
            var list = (from tbGatedot in myModels.SYS_Gatedot
                        join tbCity in myModels.BS_City on tbGatedot.CityID equals tbCity.CityID
                        select new Gatedot
                        {
                            CityID = tbGatedot.CityID,
                            CityName = tbCity.CityName,
                            Describe = tbGatedot.Describe,
                            GatedotCode = tbGatedot.GatedotCode,
                            GatedotID = tbGatedot.GatedotID,
                            GatedotName = tbGatedot.GatedotName,
                            Province = tbCity.Province,
                            WhetherSuperiorGatedota = "",
                            WhetherSuperiorGatedotb = tbGatedot.WhetherSuperiorGatedot,
                            WhetherValida = "",
                            WhetherValidb = tbGatedot.WhetherValid
                        }).ToList();
            for (int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherSuperiorGatedotb == true)
                {
                    list[i].WhetherSuperiorGatedota = "有上级门点";
                }
                else if (list[i].WhetherSuperiorGatedotb == false)
                {
                    list[i].WhetherSuperiorGatedota = "没有上级门点";
                }
            }
            for (int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherValidb == true)
                {
                    list[i].WhetherValida = "有效";
                }
                else if (list[i].WhetherValidb == false)
                {
                    list[i].WhetherValida = "无效";
                }
            }
            //把linq类型的数据listResult转化为DataTable类型数据
            DataTable dt = LINQToDataTablel(list);
            //第一步：实例化数据集`                   
            PrintReport.ReportDB dbReport = new PrintReport.ReportDB();
            //第二步：将dt的数据放入数据集的数据表中
            dbReport.Tables["Gatedott"].Merge(dt);
            //第三步：实例化报表模板
            PrintReport.gatedott rp = new PrintReport.gatedott();
            //第四步：获取报表物理文件地址     
            string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                + "Areas\\Basics\\PrintReport\\gatedott.rpt";
            //第五步：把报表文件加载到ReportDocument
            rp.Load(strRptPath);
            //第六步：设置报表数据源
            rp.SetDataSource(dbReport);
            //第七步：把ReportDocument转化为文件流
            Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
            return File(stream, "application/pdf");
        }
        // //将IEnumerable<T>类型的集合转换为DataTable类型
        public DataTable LINQToDataTablel<T>(IEnumerable<T> varlist)
        {
            //定义要返回的DataTable对象
            DataTable dtReturn = new DataTable();
            //保存列集合的属性信息数组
            PropertyInfo[] oProps = null;
            if (varlist == null) return dtReturn;//安全性检查
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
                        //得到属性的类型
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
        #region 司机信息
        /// <summary>
        /// 新增查询客户
        /// </summary>
        /// <returns></returns>
        public ActionResult selecttClientttype(int MessageID)
        {
            var list = from tbClientType in myModels.SYS_ClientType
                       join tbClient in myModels.SYS_Client on tbClientType.ClientID equals tbClient.ClientID
                       where tbClientType.ClientType.Trim() == "拖车公司" && tbClient.MessageID == MessageID
                       select new
                       {
                           id = tbClientType.ClientTypeID,
                           name = tbClient.ChineseName
                       };
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 修改查询客户
        /// </summary>
        /// <returns></returns>
        public ActionResult selecttCietype(string ClientType)
        {

            return Json("", JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 查询司机信息
        /// </summary>
        /// <returns></returns>
        public ActionResult selectChauffeur(BsgridPage bsgridPage,string ChauffeurNumber,string ChauffeurName)
        {
            var list = (from tbChauffeur in myModels.SYS_Chauffeur
                        join tbClientType in myModels.SYS_ClientType on tbChauffeur.ChauffeurID equals tbClientType.ClientTypeID
                        join tbMessage in myModels.SYS_Message on tbChauffeur.MessageID equals tbMessage.MessageID
                        select new Chauffeur
                        {
                            ChauffeurID = tbChauffeur.ChauffeurID,
                            ChauffeurNumber = tbChauffeur.ChauffeurNumber,
                            ChauffeurName = tbChauffeur.ChauffeurName,
                            Sex = tbChauffeur.Sex,
                            DeductRatio = tbChauffeur.DeductRatio,
                            ICNumber = tbChauffeur.ICNumber,
                            PhoneOne = tbChauffeur.PhoneOne,
                            WhetherStart = tbChauffeur.WhetherStart,
                            MessageID = tbChauffeur.MessageID,
                            ChineseName = tbMessage.ChineseName,
                            ClientTypeID = tbClientType.ClientTypeID,
                            ClientType = tbClientType.ClientType,
                            EnglishName = tbChauffeur.EnglishName,
                            State = tbChauffeur.State,
                            Address = tbChauffeur.Address,
                            Birthday = tbChauffeur.Birthday,
                            HealthCondition = tbChauffeur.HealthCondition,
                            IdentityCard = tbChauffeur.IdentityCard,
                            PhoneTwo = tbChauffeur.PhoneTwo,
                            HongkongPhoneOne = tbChauffeur.HongkongPhoneOne,
                            HongkongPhoneTwo = tbChauffeur.HongkongPhoneTwo,
                            FamilyPhone = tbChauffeur.FamilyPhone,
                            UrgencyRelationer = tbChauffeur.UrgencyRelationer,
                            UrgencyRelationerPhone = tbChauffeur.UrgencyRelationerPhone,
                            UrgencyRelationerSite = tbChauffeur.UrgencyRelationerSite,
                            CashDeposit = tbChauffeur.CashDeposit,
                            Salary = tbChauffeur.Salary,
                            DrivingLicence = tbChauffeur.DrivingLicence,
                            Jsznsrq = tbChauffeur.Jsznsrq,
                            Jszyxrq = tbChauffeur.Jszyxrq,
                            Hkjszh = tbChauffeur.Hkjszh,
                            Hkjszhnsrq = tbChauffeur.Hkjszhnsrq,
                            Hkjszhlyrq = tbChauffeur.Hkjszhlyrq,
                            ICNsrq = tbChauffeur.ICNsrq,
                            Hxzh = tbChauffeur.Hxzh,
                            StatusCard = tbChauffeur.StatusCard,
                            StatusCardNsrq = tbChauffeur.StatusCardNsrq,
                            SocialSecurityCardNumber = tbChauffeur.SocialSecurityCardNumber,
                            CertificateQualification = tbChauffeur.CertificateQualification,
                            CertificateQualificationNsrq=tbChauffeur.CertificateQualificationNsrq,
                            WorkLicense=tbChauffeur.WorkLicense,
                            WorkLicenseNsrq=tbChauffeur.WorkLicenseNsrq,
                            FenGSKJ=tbChauffeur.FenGSKJ,
                            Apartment=tbChauffeur.Apartment,
                            RegisteredPermanent=tbChauffeur.RegisteredPermanent,
                            HongKongPermanent=tbChauffeur.HongKongPermanent,
                            Describe=tbChauffeur.Describe
                        }).ToList();
            if (!string.IsNullOrEmpty(ChauffeurNumber))
            {
                list = list.Where(m => m.ChauffeurNumber.Contains(ChauffeurNumber.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ChauffeurName))
            {
                list = list.Where(m => m.ChauffeurName.Contains(ChauffeurName.Trim())).ToList();
            }
            int totalRow = list.Count();
            List<Chauffeur> notices = list.OrderByDescending(p => p.ChauffeurID).
              Skip(bsgridPage.GetStartIndex()).
              Take(bsgridPage.pageSize).
              ToList();
            Bsgrid<Chauffeur> bsgrid = new Bsgrid<Chauffeur>();
            bsgrid.success = true;
            bsgrid.totalRows = totalRow;
            bsgrid.curPage = bsgridPage.curPage;
            bsgrid.data = notices;
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 新增司机信息
        /// </summary>
        /// <param name="tbchauffeur"></param>
        /// <returns></returns>
        public ActionResult insertChauffeur(SYS_Chauffeur tbchauffeur,HttpPostedFileBase Image)
        {
            string strMsg = "fail";
            try
            {
                int Number = (from tbChauffeur in myModels.SYS_Chauffeur
                                    where tbChauffeur.ChauffeurNumber == tbchauffeur.ChauffeurNumber
                                    select tbChauffeur).Count();
                if (Number == 0)
                {
                    int Nmae = (from tb in myModels.SYS_Chauffeur
                                where tb.ChauffeurName == tbchauffeur.ChauffeurName
                                select tb).Count();
                    if (Nmae == 0)
                    {
                        int IdentityCard = (from tb in myModels.SYS_Chauffeur
                                            where tb.IdentityCard == tbchauffeur.IdentityCard
                                            select tb).Count();
                        if (IdentityCard == 0)
                        {
                            byte[] imgFile = null;
                            if (Image != null && Image.ContentLength > 0)
                            {
                                imgFile = new byte[Image.ContentLength];
                                Image.InputStream.Read(imgFile, 0, Image.ContentLength);
                                tbchauffeur.Picture = imgFile;
                            }
                            else
                            {
                                strMsg = "请上传相片！";
                            }
                            myModels.SYS_Chauffeur.Add(tbchauffeur);
                            if (myModels.SaveChanges() > 0)
                            {
                                strMsg = "新增成功！";
                            }
                            else
                            {
                                strMsg = "新增不成功！";
                            }
                        }
                        else
                        {
                            strMsg = "已有身份证，不能新增！";
                        }
                    }
                    else{
                        strMsg = "已有司机姓名，不能新增！";
                    }
                }
                else
                {
                    strMsg = "已有司机编号，不能新增！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 查询司机图片
        /// </summary>
        /// <returns></returns>
        public ActionResult GetChauffeurImage(int ChauffeurID)
        {
            try
            {
                var ChauffeurImg = (from tbChauffeur in myModels.SYS_Chauffeur
                                  where tbChauffeur.ChauffeurID == ChauffeurID
                                    select new
                                  {
                                      tbChauffeur.Picture
                                  }).Single();
                byte[] imageData = ChauffeurImg.Picture;
                return File(imageData, @"image/jpg");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return null;
            }
        }
        /// <summary>
        /// 修改司机信息
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateChauffeur(SYS_Chauffeur pwChauffeur, HttpPostedFileBase fileStudentImage)
        {
            string strMsg = "fail";
            try
            {
                if (myModels.SYS_Chauffeur.Where(m => m.ChauffeurNumber == pwChauffeur.ChauffeurNumber && m.ChauffeurID != pwChauffeur.ChauffeurID).Count() == 0)
                {
                    if (myModels.SYS_Chauffeur.Where(m => m.ChauffeurName == pwChauffeur.ChauffeurName && m.ChauffeurID != pwChauffeur.ChauffeurID).Count() == 0)
                    {
                        if (myModels.SYS_Chauffeur.Where(m => m.IdentityCard == pwChauffeur.IdentityCard && m.ChauffeurID != pwChauffeur.ChauffeurID).Count() == 0)
                        {
                            SYS_Chauffeur dbChauffeur = (from tbChauffeur in myModels.SYS_Chauffeur
                                                         where tbChauffeur.ChauffeurID == pwChauffeur.ChauffeurID
                                                         select tbChauffeur).Single();
                            dbChauffeur.ChauffeurName = pwChauffeur.ChauffeurName;
                            dbChauffeur.ChauffeurNumber = pwChauffeur.ChauffeurNumber;
                            dbChauffeur.Sex = pwChauffeur.Sex.Trim();
                            dbChauffeur.EnglishName = pwChauffeur.EnglishName;
                            dbChauffeur.State = pwChauffeur.State;
                            dbChauffeur.Address = pwChauffeur.Address.Trim();
                            dbChauffeur.Birthday = pwChauffeur.Birthday;
                            dbChauffeur.ClientTypeID = pwChauffeur.ClientTypeID;
                            dbChauffeur.HealthCondition = pwChauffeur.HealthCondition;
                            dbChauffeur.IdentityCard = pwChauffeur.IdentityCard;
                            dbChauffeur.PhoneOne = pwChauffeur.PhoneOne;
                            dbChauffeur.PhoneTwo = pwChauffeur.PhoneTwo;
                            dbChauffeur.HongkongPhoneOne = pwChauffeur.HongkongPhoneOne;
                            dbChauffeur.HongkongPhoneTwo = pwChauffeur.HongkongPhoneTwo;
                            dbChauffeur.FamilyPhone = pwChauffeur.FamilyPhone;
                            dbChauffeur.UrgencyRelationer = pwChauffeur.UrgencyRelationer;
                            dbChauffeur.UrgencyRelationerPhone = pwChauffeur.UrgencyRelationerPhone;
                            dbChauffeur.UrgencyRelationerSite = pwChauffeur.UrgencyRelationerSite;
                            dbChauffeur.CashDeposit = pwChauffeur.CashDeposit;
                            dbChauffeur.DeductRatio = pwChauffeur.DeductRatio;
                            dbChauffeur.Salary = pwChauffeur.Salary;
                            dbChauffeur.DrivingLicence = pwChauffeur.DrivingLicence;
                            dbChauffeur.Jsznsrq = pwChauffeur.Jsznsrq;
                            dbChauffeur.Jszyxrq = pwChauffeur.Jszyxrq;
                            dbChauffeur.Hkjszh = pwChauffeur.Hkjszh;
                            dbChauffeur.Hkjszhnsrq = pwChauffeur.Hkjszhnsrq;
                            dbChauffeur.Hkjszhlyrq = pwChauffeur.Hkjszhlyrq;
                            dbChauffeur.ICNumber = pwChauffeur.ICNumber;
                            dbChauffeur.ICNsrq = pwChauffeur.ICNsrq;
                            dbChauffeur.Hxzh = pwChauffeur.Hxzh;
                            dbChauffeur.StatusCard = pwChauffeur.StatusCard;
                            dbChauffeur.StatusCardNsrq = pwChauffeur.StatusCardNsrq;
                            dbChauffeur.SocialSecurityCardNumber = pwChauffeur.SocialSecurityCardNumber;
                            dbChauffeur.CertificateQualification = pwChauffeur.CertificateQualification;
                            dbChauffeur.CertificateQualificationNsrq = pwChauffeur.CertificateQualificationNsrq;
                            dbChauffeur.WorkLicense = pwChauffeur.WorkLicense;
                            dbChauffeur.WorkLicenseNsrq = pwChauffeur.WorkLicenseNsrq;
                            dbChauffeur.FenGSKJ = pwChauffeur.FenGSKJ;
                            dbChauffeur.WhetherStart = pwChauffeur.WhetherStart;
                            dbChauffeur.Apartment = pwChauffeur.Apartment;
                            dbChauffeur.RegisteredPermanent = pwChauffeur.RegisteredPermanent;
                            dbChauffeur.HongKongPermanent = pwChauffeur.HongKongPermanent;
                            dbChauffeur.Describe = pwChauffeur.Describe;
                            dbChauffeur.MessageID = pwChauffeur.MessageID;
                            //判断是否上传图片
                            if (fileStudentImage != null && fileStudentImage.ContentLength > 0)
                            {
                                byte[] imgFile = null;
                                imgFile = new byte[fileStudentImage.ContentLength];
                                fileStudentImage.InputStream.Read(imgFile, 0, fileStudentImage.ContentLength);
                                dbChauffeur.Picture = imgFile;
                            }
                            myModels.Entry(dbChauffeur).State = System.Data.Entity.EntityState.Modified;
                            if (myModels.SaveChanges() > 0)
                            {
                                strMsg = "修改成功！";
                            }
                        }
                        else
                        {
                            strMsg = "已有身份证，无法修改！";
                        }
                    }
                    else
                    {
                        strMsg = "已有司机姓名，无法修改！";
                    }
                }
                else
                {
                    strMsg = "已有司机编号，无法修改！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 回填司机信息
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectChauffeurById(int ChauffeurID)
        {
            try
            {
                var list = (from tbChauffeur in myModels.SYS_Chauffeur
                            join tbClientType in myModels.SYS_ClientType on tbChauffeur.ChauffeurID equals tbClientType.ClientTypeID
                            join tbMessage in myModels.SYS_Message on tbChauffeur.MessageID equals tbMessage.MessageID
                            where tbChauffeur.ChauffeurID == ChauffeurID
                            select new Chauffeur
                            {
                                ChauffeurID = tbChauffeur.ChauffeurID,
                                ChauffeurNumber = tbChauffeur.ChauffeurNumber,
                                ChauffeurName = tbChauffeur.ChauffeurName,
                                Sex = tbChauffeur.Sex.Trim(),
                                DeductRatio = tbChauffeur.DeductRatio,
                                ICNumber = tbChauffeur.ICNumber,
                                PhoneOne = tbChauffeur.PhoneOne,
                                WhetherStart = tbChauffeur.WhetherStart,
                                MessageID = tbChauffeur.MessageID,
                                ChineseName = tbMessage.ChineseName,
                                ClientTypeID = tbClientType.ClientTypeID,
                                ClientType = tbClientType.ClientType,
                                EnglishName = tbChauffeur.EnglishName,
                                State = tbChauffeur.State.Trim(),
                                Address = tbChauffeur.Address.Trim(),
                                Birthday = tbChauffeur.Birthday,
                                HealthCondition = tbChauffeur.HealthCondition.Trim(),
                                IdentityCard = tbChauffeur.IdentityCard,
                                PhoneTwo = tbChauffeur.PhoneTwo,
                                HongkongPhoneOne = tbChauffeur.HongkongPhoneOne,
                                HongkongPhoneTwo = tbChauffeur.HongkongPhoneTwo,
                                FamilyPhone = tbChauffeur.FamilyPhone,
                                UrgencyRelationer = tbChauffeur.UrgencyRelationer,
                                UrgencyRelationerPhone = tbChauffeur.UrgencyRelationerPhone,
                                UrgencyRelationerSite = tbChauffeur.UrgencyRelationerSite,
                                CashDeposit = tbChauffeur.CashDeposit,
                                Salary = tbChauffeur.Salary,
                                DrivingLicence = tbChauffeur.DrivingLicence,
                                Jsznsrq = tbChauffeur.Jsznsrq,
                                Jszyxrq = tbChauffeur.Jszyxrq,
                                Hkjszh = tbChauffeur.Hkjszh,
                                Hkjszhnsrq = tbChauffeur.Hkjszhnsrq,
                                Hkjszhlyrq = tbChauffeur.Hkjszhlyrq,
                                ICNsrq = tbChauffeur.ICNsrq,
                                Hxzh = tbChauffeur.Hxzh,
                                StatusCard = tbChauffeur.StatusCard,
                                StatusCardNsrq = tbChauffeur.StatusCardNsrq,
                                SocialSecurityCardNumber = tbChauffeur.SocialSecurityCardNumber,
                                CertificateQualification = tbChauffeur.CertificateQualification,
                                CertificateQualificationNsrq = tbChauffeur.CertificateQualificationNsrq,
                                WorkLicense = tbChauffeur.WorkLicense,
                                WorkLicenseNsrq = tbChauffeur.WorkLicenseNsrq,
                                FenGSKJ = tbChauffeur.FenGSKJ,
                                Apartment = tbChauffeur.Apartment,
                                RegisteredPermanent = tbChauffeur.RegisteredPermanent,
                                HongKongPermanent = tbChauffeur.HongKongPermanent,
                                Describe = tbChauffeur.Describe
                            }).Single();
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return null;
            }
        }
        /// <summary>
        /// 查询司机照片
        /// </summary>
        /// <returns></returns>
        public ActionResult ChauffeurImage(int ChauffeurID)
        {
            try
            {
                var Img = (from tbChauffeur in myModels.SYS_Chauffeur
                                  where tbChauffeur.ChauffeurID == ChauffeurID
                                  select new
                                  {
                                      tbChauffeur.Picture
                                  }).Single();
                byte[] imageData = Img.Picture;
                return File(imageData, @"image/jpg");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return null;
            }
        }
        /// <summary>
        /// 删除司机信息
        /// </summary>
        /// <returns></returns>
        public ActionResult DelectChauffeur(int ChauffeurID)
        {
            string strMsg = "fail";
            try
            {
                var listChauffeur = myModels.SYS_Chauffeur.Where(m => m.ChauffeurID == ChauffeurID).Single();
                myModels.SYS_Chauffeur.Remove(listChauffeur);
                if (myModels.SaveChanges() > 0)
                {
                    strMsg = "true";
                }
            }
            catch (Exception e)
            {
                strMsg = "fail";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 导出司机信息
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportToExcelChauffeur(string ChauffeurNumber, string ChauffeurName)
        {
            var list = (from tbChauffeur in myModels.SYS_Chauffeur
                        join tbClientType in myModels.SYS_ClientType on tbChauffeur.ChauffeurID equals tbClientType.ClientTypeID
                        join tbMessage in myModels.SYS_Message on tbChauffeur.MessageID equals tbMessage.MessageID
                        select new Chauffeur
                        {
                            ChauffeurID = tbChauffeur.ChauffeurID,
                            ChauffeurNumber = tbChauffeur.ChauffeurNumber,
                            ChauffeurName = tbChauffeur.ChauffeurName,
                            Sex = tbChauffeur.Sex,
                            DeductRatio = tbChauffeur.DeductRatio,
                            ICNumber = tbChauffeur.ICNumber,
                            PhoneOne = tbChauffeur.PhoneOne,
                            MessageID = tbChauffeur.MessageID,
                            ChineseName = tbMessage.ChineseName,
                            ClientTypeID = tbClientType.ClientTypeID,
                            ClientType = tbClientType.ClientType,
                            EnglishName = tbChauffeur.EnglishName,
                            State = tbChauffeur.State,
                            Address = tbChauffeur.Address,
                            Birthday = tbChauffeur.Birthday,
                            HealthCondition = tbChauffeur.HealthCondition,
                            IdentityCard = tbChauffeur.IdentityCard,
                            PhoneTwo = tbChauffeur.PhoneTwo,
                            HongkongPhoneOne = tbChauffeur.HongkongPhoneOne,
                            HongkongPhoneTwo = tbChauffeur.HongkongPhoneTwo,
                            FamilyPhone = tbChauffeur.FamilyPhone,
                            UrgencyRelationer = tbChauffeur.UrgencyRelationer,
                            UrgencyRelationerPhone = tbChauffeur.UrgencyRelationerPhone,
                            UrgencyRelationerSite = tbChauffeur.UrgencyRelationerSite,
                            CashDeposit = tbChauffeur.CashDeposit,
                            Salary = tbChauffeur.Salary,
                            DrivingLicence = tbChauffeur.DrivingLicence,
                            Jsznsrq = tbChauffeur.Jsznsrq,
                            Jszyxrq = tbChauffeur.Jszyxrq,
                            Hkjszh = tbChauffeur.Hkjszh,
                            Hkjszhnsrq = tbChauffeur.Hkjszhnsrq,
                            Hkjszhlyrq = tbChauffeur.Hkjszhlyrq,
                            ICNsrq = tbChauffeur.ICNsrq,
                            Hxzh = tbChauffeur.Hxzh,
                            StatusCard = tbChauffeur.StatusCard,
                            StatusCardNsrq = tbChauffeur.StatusCardNsrq,
                            SocialSecurityCardNumber = tbChauffeur.SocialSecurityCardNumber,
                            CertificateQualification = tbChauffeur.CertificateQualification,
                            CertificateQualificationNsrq = tbChauffeur.CertificateQualificationNsrq,
                            WorkLicense = tbChauffeur.WorkLicense,
                            WorkLicenseNsrq = tbChauffeur.WorkLicenseNsrq,
                            FenGSKJ = tbChauffeur.FenGSKJ,
                            Apartment = tbChauffeur.Apartment,
                            RegisteredPermanent = tbChauffeur.RegisteredPermanent,
                            HongKongPermanent = tbChauffeur.HongKongPermanent,
                            Describe = tbChauffeur.Describe,
                            WhetherStarta="",
                            WhetherStartb = tbChauffeur.WhetherStart,
                        }).ToList();
            if (!string.IsNullOrEmpty(ChauffeurNumber))
            {
                list = list.Where(m => m.ChauffeurNumber.Contains(ChauffeurNumber.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ChauffeurName))
            {
                list = list.Where(m => m.ChauffeurName.Contains(ChauffeurName.Trim())).ToList();
            }
            for (int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherStartb == true)
                {
                    list[i].WhetherStarta = "启用";
                }
                else if (list[i].WhetherStartb == false)
                {
                    list[i].WhetherStarta = "不启用";
                }
            }
            //查询数据
            List<Chauffeur> listExaminee = list.ToList();
            //二：代码创建一个Excel表格（这里称为工作簿）
            //创建Excel文件的对象 工作簿(调用NPOI文件)
            HSSFWorkbook excelBook = new HSSFWorkbook();
            //创建Excel工作表 Sheet
            ISheet sheet1 = excelBook.CreateSheet("司机信息");
            //给Sheet添加第一行的头部标题
            IRow row1 = sheet1.CreateRow(0);
            //给标题的每一个单元格赋值
            row1.CreateCell(0).SetCellValue("司机姓名");
            row1.CreateCell(1).SetCellValue("司机编号");
            row1.CreateCell(2).SetCellValue("性别");
            row1.CreateCell(3).SetCellValue("状态");
            row1.CreateCell(4).SetCellValue("籍贯");
            row1.CreateCell(5).SetCellValue("健康状况");
            row1.CreateCell(6).SetCellValue("身份证");
            row1.CreateCell(7).SetCellValue("手机");
            row1.CreateCell(8).SetCellValue("提成比例%");
            row1.CreateCell(9).SetCellValue("是否启用");
            //添加数据行：将表格数据逐步写入sheet1各个行中（也就是给每一个单元格赋值）
            for (int i = 0; i < listExaminee.Count; i++)
            {
                //创建行
                IRow rowTemp = sheet1.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listExaminee[i].ChauffeurName);
                rowTemp.CreateCell(1).SetCellValue(listExaminee[i].ChauffeurNumber);
                rowTemp.CreateCell(2).SetCellValue(listExaminee[i].Sex);
                rowTemp.CreateCell(3).SetCellValue(listExaminee[i].State);
                rowTemp.CreateCell(4).SetCellValue(listExaminee[i].Address);
                rowTemp.CreateCell(5).SetCellValue(listExaminee[i].HealthCondition);
                rowTemp.CreateCell(6).SetCellValue(listExaminee[i].IdentityCard);
                rowTemp.CreateCell(7).SetCellValue(listExaminee[i].PhoneOne);
                rowTemp.CreateCell(8).SetCellValue(listExaminee[i].DeductRatio.ToString());
                rowTemp.CreateCell(9).SetCellValue(listExaminee[i].WhetherStarta);
            }
            //输出的文件名称
            string fileName = "司机信息" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
            //把Excel转为流，输出
            //创建文件流
            System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
            //将工作薄写入文件流
            excelBook.Write(bookStream);
            //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
            bookStream.Seek(0, System.IO.SeekOrigin.Begin);
            //Stream对象,文件类型,文件名称
            return File(bookStream, "application/vnd.ms-excel", fileName);
        }
        /// <summary>
        /// 司机水晶报表
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievementChauffeur(string ChauffeurNumber, string ChauffeurName)
        {
            var list = (from tbChauffeur in myModels.SYS_Chauffeur
                        join tbClientType in myModels.SYS_ClientType on tbChauffeur.ChauffeurID equals tbClientType.ClientTypeID
                        join tbMessage in myModels.SYS_Message on tbChauffeur.MessageID equals tbMessage.MessageID
                        select new Chauffeur
                        {
                            ChauffeurID = tbChauffeur.ChauffeurID,
                            ChauffeurNumber = tbChauffeur.ChauffeurNumber,
                            ChauffeurName = tbChauffeur.ChauffeurName,
                            Sex = tbChauffeur.Sex,
                            DeductRatio = tbChauffeur.DeductRatio,
                            ICNumber = tbChauffeur.ICNumber,
                            PhoneOne = tbChauffeur.PhoneOne,
                            MessageID = tbChauffeur.MessageID,
                            ChineseName = tbMessage.ChineseName,
                            ClientTypeID = tbClientType.ClientTypeID,
                            ClientType = tbClientType.ClientType,
                            EnglishName = tbChauffeur.EnglishName,
                            State = tbChauffeur.State,
                            Address = tbChauffeur.Address,
                            Birthday = tbChauffeur.Birthday,
                            HealthCondition = tbChauffeur.HealthCondition,
                            IdentityCard = tbChauffeur.IdentityCard,
                            PhoneTwo = tbChauffeur.PhoneTwo,
                            HongkongPhoneOne = tbChauffeur.HongkongPhoneOne,
                            HongkongPhoneTwo = tbChauffeur.HongkongPhoneTwo,
                            FamilyPhone = tbChauffeur.FamilyPhone,
                            UrgencyRelationer = tbChauffeur.UrgencyRelationer,
                            UrgencyRelationerPhone = tbChauffeur.UrgencyRelationerPhone,
                            UrgencyRelationerSite = tbChauffeur.UrgencyRelationerSite,
                            CashDeposit = tbChauffeur.CashDeposit,
                            Salary = tbChauffeur.Salary,
                            DrivingLicence = tbChauffeur.DrivingLicence,
                            Jsznsrq = tbChauffeur.Jsznsrq,
                            Jszyxrq = tbChauffeur.Jszyxrq,
                            Hkjszh = tbChauffeur.Hkjszh,
                            Hkjszhnsrq = tbChauffeur.Hkjszhnsrq,
                            Hkjszhlyrq = tbChauffeur.Hkjszhlyrq,
                            ICNsrq = tbChauffeur.ICNsrq,
                            Hxzh = tbChauffeur.Hxzh,
                            StatusCard = tbChauffeur.StatusCard,
                            StatusCardNsrq = tbChauffeur.StatusCardNsrq,
                            SocialSecurityCardNumber = tbChauffeur.SocialSecurityCardNumber,
                            CertificateQualification = tbChauffeur.CertificateQualification,
                            CertificateQualificationNsrq = tbChauffeur.CertificateQualificationNsrq,
                            WorkLicense = tbChauffeur.WorkLicense,
                            WorkLicenseNsrq = tbChauffeur.WorkLicenseNsrq,
                            FenGSKJ = tbChauffeur.FenGSKJ,
                            Apartment = tbChauffeur.Apartment,
                            RegisteredPermanent = tbChauffeur.RegisteredPermanent,
                            HongKongPermanent = tbChauffeur.HongKongPermanent,
                            Describe = tbChauffeur.Describe,
                            WhetherStarta = "",
                            WhetherStartb = tbChauffeur.WhetherStart,
                        }).ToList();
            if (!string.IsNullOrEmpty(ChauffeurNumber))
            {
                list = list.Where(m => m.ChauffeurNumber.Contains(ChauffeurNumber.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ChauffeurName))
            {
                list = list.Where(m => m.ChauffeurName.Contains(ChauffeurName.Trim())).ToList();
            }
            for (int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherStartb == true)
                {
                    list[i].WhetherStarta = "√";
                }
                else if (list[i].WhetherStartb == false)
                {
                    list[i].WhetherStarta = "×";
                }
            }
            //把linq类型的数据listResult转化为DataTable类型数据
            DataTable dt = LINQToDataTableChauffeur(list);
            //第一步：实例化数据集`
            PrintReport.ReportDB dbReport = new PrintReport.ReportDB();
            //第二步：将dt的数据放入数据集的数据表中
            dbReport.Tables["Chauffeur"].Merge(dt);
            //第三步：实例化报表模板
            PrintReport.chauffeur rp = new PrintReport.chauffeur();
            //第四步：获取报表物理文件地址
            string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                + "Areas\\Basics\\PrintReport\\chauffeur.rpt";
            //第五步：把报表文件加载到ReportDocument
            rp.Load(strRptPath);
            //第六步：设置报表数据源
            rp.SetDataSource(dbReport);
            //第七步：把ReportDocument转化为文件流
            Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
            return File(stream, "application/pdf");
        }
        // //将IEnumerable<T>类型的集合转换为DataTable类型
        public DataTable LINQToDataTableChauffeur<T>(IEnumerable<T> varlist)
        {
            //定义要返回的DataTable对象
            DataTable dtReturn = new DataTable();
            //保存列集合的属性信息数组
            PropertyInfo[] oProps = null;
            if (varlist == null) return dtReturn;//安全性检查
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
                        //得到属性的类型
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
        #region 司机本▓
        /// <summary>
        /// 查询司机本
        /// </summary>
        /// <returns></returns>
        public ActionResult selectDriverBook(BsgridPage bsgridPage,int? ChauffeurID,string DriverBook,string CustomsNumber)
        {
            var list = (from tbDriverBook in myModels.SYS_DriverBook
                        join tbChauffeur in myModels.SYS_Chauffeur on tbDriverBook.ChauffeurID equals tbChauffeur.ChauffeurID
                        select new DriverBook
                        {
                            ChauffeurID = tbDriverBook.ChauffeurID,
                            ChauffeurName=tbChauffeur.ChauffeurName,
                            DriverBookID = tbDriverBook.DriverBookID,
                            DriverBook = tbDriverBook.DriverBook,
                            CustomsNumber = tbDriverBook.CustomsNumber,
                            BusinessLicenseNumber = tbDriverBook.BusinessLicenseNumber,
                            AccurateLoadNumber = tbDriverBook.AccurateLoadNumber,
                            IssuingAuthority = tbDriverBook.IssuingAuthority,
                            TermValidity = tbDriverBook.TermValidity,
                            timer = tbDriverBook.TermValidity.ToString(),
                            ExpirationDate = tbDriverBook.ExpirationDate,
                            ExpirationDatetimer = tbDriverBook.ExpirationDate.ToString(),
                            Remark = tbDriverBook.Remark
                        }).ToList();
            if (ChauffeurID > 0)
            {
                list = list.Where(m => m.ChauffeurID == ChauffeurID).ToList();
            }
            if (!string.IsNullOrEmpty(DriverBook))
            {
                list = list.Where(m => m.DriverBook.Contains(DriverBook.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(CustomsNumber))
            {
                list = list.Where(m => m.CustomsNumber.Contains(CustomsNumber.Trim())).ToList();
            }
            int totalRow = list.Count();
            List<DriverBook> notices = list.OrderByDescending(p => p.DriverBookID).
              Skip(bsgridPage.GetStartIndex()).
              Take(bsgridPage.pageSize).
              ToList();
            Bsgrid<DriverBook> bsgrid = new Bsgrid<DriverBook>();
            bsgrid.success = true;
            bsgrid.totalRows = totalRow;
            bsgrid.curPage = bsgridPage.curPage;
            bsgrid.data = notices;
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 新增司机本
        /// </summary>
        /// <returns></returns>
        public ActionResult InsertDriverBook(SYS_DriverBook tbDriverBook)
        {
            string strMsg = "fail";
            try
            {
                var Book = (from tb in myModels.SYS_DriverBook
                            where tb.DriverBook == tbDriverBook.DriverBook
                            select tb).Count();
                if (Book == 0)
                {
                    var AccurateLoadNumber = (from tb in myModels.SYS_DriverBook
                                              where tb.AccurateLoadNumber == tbDriverBook.AccurateLoadNumber
                                              select tb).Count();
                    if (AccurateLoadNumber == 0)
                    {
                        var BusinessLicenseNumber = (from tb in myModels.SYS_DriverBook
                                                     where tb.BusinessLicenseNumber == tbDriverBook.BusinessLicenseNumber
                                                     select tb).Count();
                        if (BusinessLicenseNumber == 0)
                        {
                            myModels.SYS_DriverBook.Add(tbDriverBook);
                            if (myModels.SaveChanges() > 0)
                            {
                                strMsg = "数据新增成功！";
                            }
                        }
                        else
                        {
                            strMsg = "已有营业执照号，无法新增！";
                        }
                    }
                    else {
                        strMsg = "已有准载证号，无法新增！";
                    }
                }
                else
                {
                    strMsg = "已有编号，无法新增！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 修改司机本
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateDriverBook(SYS_DriverBook pwDriverBook)
        {
            string strMsg = "fail";
            try
            {
                if (myModels.SYS_DriverBook.Where(m => m.DriverBook == pwDriverBook.DriverBook && m.DriverBookID != pwDriverBook.DriverBookID).Count() == 0)
                {
                    if (myModels.SYS_DriverBook.Where(m => m.AccurateLoadNumber == pwDriverBook.AccurateLoadNumber && m.DriverBookID != pwDriverBook.DriverBookID).Count() == 0)
                    {
                        if (myModels.SYS_DriverBook.Where(m => m.BusinessLicenseNumber == pwDriverBook.BusinessLicenseNumber && m.DriverBookID != pwDriverBook.DriverBookID).Count() == 0)
                        {
                            SYS_DriverBook dbDriverBook = (from tbDriverBook in myModels.SYS_DriverBook
                                                           where tbDriverBook.DriverBookID == pwDriverBook.DriverBookID
                                                           select tbDriverBook).Single();
                            dbDriverBook.ChauffeurID = pwDriverBook.ChauffeurID;
                            dbDriverBook.AccurateLoadNumber = pwDriverBook.AccurateLoadNumber;
                            dbDriverBook.BusinessLicenseNumber = pwDriverBook.BusinessLicenseNumber;
                            dbDriverBook.CustomsNumber = pwDriverBook.CustomsNumber;
                            dbDriverBook.DriverBook = pwDriverBook.DriverBook;
                            dbDriverBook.IssuingAuthority = pwDriverBook.IssuingAuthority;
                            dbDriverBook.Remark = pwDriverBook.Remark;
                            dbDriverBook.TermValidity = pwDriverBook.TermValidity;
                            dbDriverBook.ExpirationDate = pwDriverBook.ExpirationDate;
                            myModels.Entry(dbDriverBook).State = System.Data.Entity.EntityState.Modified;
                            if (myModels.SaveChanges() > 0)
                            {
                                strMsg = "修改成功！";
                            }
                        }
                        else
                        {
                            strMsg = "已有营业执照号，无法修改！";
                        }
                    }
                    else
                    {
                        strMsg = "已有准载证号，无法修改！";
                    }
                }
                else
                {
                    strMsg = "已有编号，无法修改！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 回填司机本
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectDriverBookById(int DriverBookID)
        {
            try
            {
                var list = (from tbDriverBook in myModels.SYS_DriverBook
                            join tbChauffeur in myModels.SYS_Chauffeur on tbDriverBook.ChauffeurID equals tbChauffeur.ChauffeurID
                            where tbDriverBook.DriverBookID == DriverBookID
                            select new DriverBook
                            {
                                ChauffeurID = tbDriverBook.ChauffeurID,
                                ChauffeurName=tbChauffeur.ChauffeurName,
                                DriverBookID = tbDriverBook.DriverBookID,
                                DriverBook = tbDriverBook.DriverBook,
                                CustomsNumber = tbDriverBook.CustomsNumber,
                                BusinessLicenseNumber = tbDriverBook.BusinessLicenseNumber,
                                AccurateLoadNumber = tbDriverBook.AccurateLoadNumber,
                                IssuingAuthority = tbDriverBook.IssuingAuthority,
                                TermValidity = tbDriverBook.TermValidity,
                                timer = tbDriverBook.TermValidity.ToString(),
                                ExpirationDate = tbDriverBook.ExpirationDate,
                                ExpirationDatetimer = tbDriverBook.ExpirationDate.ToString(),
                                Remark = tbDriverBook.Remark
                            }).Single();
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return null;
            }
        }
        /// <summary>
        /// 删除司机本
        /// </summary>
        /// <returns></returns>
        public ActionResult DelectDriverBook(int DriverBookID)
        {
            string strMsg = "fail";
            try
            {
                var listDriverBook = myModels.SYS_DriverBook.Where(m => m.DriverBookID == DriverBookID).Single();
                myModels.SYS_DriverBook.Remove(listDriverBook);
                if (myModels.SaveChanges() > 0)
                {
                    strMsg = "true";
                }
            }
            catch (Exception e)
            {
                strMsg = "fail";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 导出司机本
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportToExcelDriverBook(int ChauffeurID, string DriverBook, string CustomsNumber)
        {
            var list = (from tbDriverBook in myModels.SYS_DriverBook
                        join tbChauffeur in myModels.SYS_Chauffeur on tbDriverBook.ChauffeurID equals tbChauffeur.ChauffeurID
                        select new DriverBook
                        {
                            ChauffeurID = tbDriverBook.ChauffeurID,
                            ChauffeurName = tbChauffeur.ChauffeurName,
                            DriverBookID = tbDriverBook.DriverBookID,
                            DriverBook = tbDriverBook.DriverBook,
                            CustomsNumber = tbDriverBook.CustomsNumber,
                            BusinessLicenseNumber = tbDriverBook.BusinessLicenseNumber,
                            AccurateLoadNumber = tbDriverBook.AccurateLoadNumber,
                            IssuingAuthority = tbDriverBook.IssuingAuthority,
                            TermValidity = tbDriverBook.TermValidity,
                            ExpirationDate = tbDriverBook.ExpirationDate,
                            Remark = tbDriverBook.Remark
                        }).ToList();
            if (ChauffeurID > 0)
            {
                list = list.Where(m => m.ChauffeurID == ChauffeurID).ToList();
            }
            if (!string.IsNullOrEmpty(DriverBook))
            {
                list = list.Where(m => m.DriverBook.Contains(DriverBook.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(CustomsNumber))
            {
                list = list.Where(m => m.CustomsNumber.Contains(CustomsNumber.Trim())).ToList();
            }
            //查询数据
            List<DriverBook> listExaminee = list.ToList();
            //二：代码创建一个Excel表格（这里称为工作簿）
            //创建Excel文件的对象 工作簿(调用NPOI文件)
            HSSFWorkbook excelBook = new HSSFWorkbook();
            //创建Excel工作表 Sheet
            ISheet sheet1 = excelBook.CreateSheet("司机本");
            //给Sheet添加第一行的头部标题
            IRow row1 = sheet1.CreateRow(0);
            //给标题的每一个单元格赋值
            row1.CreateCell(0).SetCellValue("所属司机");
            row1.CreateCell(1).SetCellValue("司机本编号");
            row1.CreateCell(2).SetCellValue("海关编号");
            row1.CreateCell(3).SetCellValue("营业执照号");
            row1.CreateCell(4).SetCellValue("核发机关");
            row1.CreateCell(5).SetCellValue("准载证号");
            row1.CreateCell(6).SetCellValue("有效期限");
            row1.CreateCell(7).SetCellValue("截止日期");
            row1.CreateCell(8).SetCellValue("备注");
            //添加数据行：将表格数据逐步写入sheet1各个行中（也就是给每一个单元格赋值）
            for (int i = 0; i < listExaminee.Count; i++)
            {
                //创建行
                IRow rowTemp = sheet1.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listExaminee[i].ChauffeurName);
                rowTemp.CreateCell(1).SetCellValue(listExaminee[i].DriverBook);
                rowTemp.CreateCell(2).SetCellValue(listExaminee[i].CustomsNumber);
                rowTemp.CreateCell(3).SetCellValue(listExaminee[i].BusinessLicenseNumber);
                rowTemp.CreateCell(4).SetCellValue(listExaminee[i].IssuingAuthority);
                rowTemp.CreateCell(5).SetCellValue(listExaminee[i].AccurateLoadNumber);
                rowTemp.CreateCell(6).SetCellValue(listExaminee[i].TermValidity.ToString());
                rowTemp.CreateCell(7).SetCellValue(listExaminee[i].ExpirationDate.ToString());
                rowTemp.CreateCell(8).SetCellValue(listExaminee[i].Remark);
            }
            //输出的文件名称
            string fileName = "司机本" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
            //把Excel转为流，输出
            //创建文件流
            System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
            //将工作薄写入文件流
            excelBook.Write(bookStream);
            //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
            bookStream.Seek(0, System.IO.SeekOrigin.Begin);
            //Stream对象,文件类型,文件名称
            return File(bookStream, "application/vnd.ms-excel", fileName);
        }
        /// <summary>
        /// 司机本报表
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievementDriverBook(int ChauffeurID, string DriverBook, string CustomsNumber)
        {
            var list = (from tbDriverBook in myModels.SYS_DriverBook
                        join tbChauffeur in myModels.SYS_Chauffeur on tbDriverBook.ChauffeurID equals tbChauffeur.ChauffeurID
                        select new DriverBook
                        {
                            ChauffeurID = tbDriverBook.ChauffeurID,
                            ChauffeurName = tbChauffeur.ChauffeurName,
                            DriverBookID = tbDriverBook.DriverBookID,
                            DriverBook = tbDriverBook.DriverBook,
                            CustomsNumber = tbDriverBook.CustomsNumber,
                            BusinessLicenseNumber = tbDriverBook.BusinessLicenseNumber,
                            AccurateLoadNumber = tbDriverBook.AccurateLoadNumber,
                            IssuingAuthority = tbDriverBook.IssuingAuthority,
                            TermValidity = tbDriverBook.TermValidity,
                            ExpirationDate = tbDriverBook.ExpirationDate,
                            Remark = tbDriverBook.Remark
                        }).ToList();
            if (ChauffeurID > 0)
            {
                list = list.Where(m => m.ChauffeurID == ChauffeurID).ToList();
            }
            if (!string.IsNullOrEmpty(DriverBook))
            {
                list = list.Where(m => m.DriverBook.Contains(DriverBook.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(CustomsNumber))
            {
                list = list.Where(m => m.CustomsNumber.Contains(CustomsNumber.Trim())).ToList();
            }
            //把linq类型的数据listResult转化为DataTable类型数据
            DataTable dt = LINQToDataTableDriverBook(list);
            //第一步：实例化数据集`
            PrintReport.ReportDB dbReport = new PrintReport.ReportDB();
            //第二步：将dt的数据放入数据集的数据表中
            dbReport.Tables["DriverBook"].Merge(dt);
            //第三步：实例化报表模板
            PrintReport.driverBook rp = new PrintReport.driverBook();
            //第四步：获取报表物理文件地址
            string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                + "Areas\\Basics\\PrintReport\\driverBook.rpt";
            //第五步：把报表文件加载到ReportDocument
            rp.Load(strRptPath);
            //第六步：设置报表数据源
            rp.SetDataSource(dbReport);
            //第七步：把ReportDocument转化为文件流
            Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
            return File(stream, "application/pdf");
        }
        // //将IEnumerable<T>类型的集合转换为DataTable类型
        public DataTable LINQToDataTableDriverBook<T>(IEnumerable<T> varlist)
        {
            //定义要返回的DataTable对象
            DataTable dtReturn = new DataTable();
            //保存列集合的属性信息数组
            PropertyInfo[] oProps = null;
            if (varlist == null) return dtReturn;//安全性检查
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
                        //得到属性的类型
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
        #region 车辆信息
        /// <summary>
        /// 查询车辆信息
        /// </summary>
        /// <returns></returns>
        public ActionResult selectVehicleInformation(BsgridPage bsgridPage, int? ChauffeurID, short? MessageID,int? BracketID,string VehicleCode,string VehicleClass,string VehicleType)
        {
            var list = (from tbVehicleInformation in myModels.SYS_VehicleInformation
                        join tbChauffeur in myModels.SYS_Chauffeur on tbVehicleInformation.ChauffeurID equals tbChauffeur.ChauffeurID
                        join tbMessage in myModels.SYS_Message on tbVehicleInformation.MessageID equals tbMessage.MessageID
                        join tbBracket in myModels.SYS_Bracket on tbVehicleInformation.BracketID equals tbBracket.BracketID
                        select new VehicleInformation
                        {
                            VehicleInformationID = tbVehicleInformation.VehicleInformationID,
                            ChauffeurID = tbVehicleInformation.ChauffeurID,
                            ChauffeurName=tbChauffeur.ChauffeurName,
                            MessageID = tbVehicleInformation.MessageID,
                            ChineseName=tbMessage.ChineseName,
                            BracketID = tbVehicleInformation.BracketID,
                            BracketCode=tbBracket.BracketCode,
                            VehicleCode = tbVehicleInformation.VehicleCode,
                            PlateNumbers = tbVehicleInformation.PlateNumbers,
                            HongKongPlateNumber = tbVehicleInformation.HongKongPlateNumber,
                            VehicleClass = tbVehicleInformation.VehicleClass,
                            MotorcycleType = tbVehicleInformation.MotorcycleType,
                            VehicleType = tbVehicleInformation.VehicleType,
                            Usualplace = tbVehicleInformation.Usualplace
                        }).ToList();
            if (ChauffeurID > 0)
            {
                list = list.Where(m => m.ChauffeurID == ChauffeurID).ToList();
            }
            if (MessageID > 0)
            {
                list = list.Where(m => m.MessageID == MessageID).ToList();
            }
            if (BracketID > 0)
            {
                list = list.Where(m => m.BracketID == BracketID).ToList();
            }
            if (!string.IsNullOrEmpty(VehicleCode))
            {
                list = list.Where(m => m.VehicleCode.Contains(VehicleCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(VehicleClass))
            {
                list = list.Where(m => m.VehicleClass.Contains(VehicleClass.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(VehicleType))
            {
                list = list.Where(m => m.VehicleType.Contains(VehicleType.Trim())).ToList();
            }
            int totalRow = list.Count();
            List<VehicleInformation> notices = list.OrderByDescending(p => p.VehicleInformationID).
              Skip(bsgridPage.GetStartIndex()).
              Take(bsgridPage.pageSize).
              ToList();
            Bsgrid<VehicleInformation> bsgrid = new Bsgrid<VehicleInformation>();
            bsgrid.success = true;
            bsgrid.totalRows = totalRow;
            bsgrid.curPage = bsgridPage.curPage;
            bsgrid.data = notices;
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 新增车辆信息
        /// </summary>
        /// <returns></returns>
        public ActionResult InsertVehicleInformation(SYS_VehicleInformation tbVehicleInformation)
        {
            string strMsg = "fail";
            try
            {
                var list = (from tbVehicle in myModels.SYS_VehicleInformation
                            where tbVehicle.VehicleCode == tbVehicleInformation.VehicleCode
                            select tbVehicle).Count();
                if (list == 0)
                {
                    var liste = (from tbveh in myModels.SYS_VehicleInformation
                                 where tbveh.PlateNumbers == tbVehicleInformation.PlateNumbers
                                 select tbveh).Count();
                    if (liste == 0)
                    {
                        myModels.SYS_VehicleInformation.Add(tbVehicleInformation);
                        if (myModels.SaveChanges() > 0)
                        {
                            strMsg = "数据新增成功！";
                        }
                    }
                    else
                    {
                        strMsg = "已有车牌号，无法新增！";
                    }
                }
                else
                {
                    strMsg = "已有车辆代码，无法新增！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 修改车辆信息
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateVehicleInformation(SYS_VehicleInformation pwVehicleInformation)
        {
            string strMsg = "fail";
            try
            {
                if (myModels.SYS_VehicleInformation.Where(m => m.VehicleCode == pwVehicleInformation.VehicleCode && m.VehicleInformationID != pwVehicleInformation.VehicleInformationID).Count() == 0)
                {
                    if (myModels.SYS_VehicleInformation.Where(m => m.PlateNumbers == pwVehicleInformation.PlateNumbers && m.VehicleInformationID != pwVehicleInformation.VehicleInformationID).Count() == 0)
                    {
                        SYS_VehicleInformation dbVehicleInformation = (from tbVehicleInformation in myModels.SYS_VehicleInformation
                                                                       where tbVehicleInformation.VehicleInformationID == pwVehicleInformation.VehicleInformationID
                                                                       select tbVehicleInformation).Single();
                        dbVehicleInformation.ChauffeurID = pwVehicleInformation.ChauffeurID;
                        dbVehicleInformation.BracketID = pwVehicleInformation.BracketID;
                        dbVehicleInformation.MessageID = pwVehicleInformation.MessageID;
                        dbVehicleInformation.HongKongPlateNumber = pwVehicleInformation.HongKongPlateNumber;
                        dbVehicleInformation.MotorcycleType = pwVehicleInformation.MotorcycleType;
                        dbVehicleInformation.PlateNumbers = pwVehicleInformation.PlateNumbers;
                        dbVehicleInformation.VehicleClass = pwVehicleInformation.VehicleClass;
                        dbVehicleInformation.VehicleCode = pwVehicleInformation.VehicleCode;
                        dbVehicleInformation.VehicleType = pwVehicleInformation.VehicleType;
                        dbVehicleInformation.Usualplace = pwVehicleInformation.Usualplace;
                        myModels.Entry(dbVehicleInformation).State = System.Data.Entity.EntityState.Modified;
                        if (myModels.SaveChanges() > 0)
                        {
                            strMsg = "修改成功！";
                        }
                    }
                    else
                    {
                        strMsg = "已有车牌号，无法修改！";
                    }
                }
                else
                {
                    strMsg = "已有车辆代码，无法修改！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 回填车辆信息
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectVehicleInformationById(int VehicleInformationID)
        {
            try
            {
                var list = (from tbVehicleInformation in myModels.SYS_VehicleInformation
                            join tbChauffeur in myModels.SYS_Chauffeur on tbVehicleInformation.ChauffeurID equals tbChauffeur.ChauffeurID
                            join tbMessage in myModels.SYS_Message on tbVehicleInformation.MessageID equals tbMessage.MessageID
                            join tbBracket in myModels.SYS_Bracket on tbVehicleInformation.BracketID equals tbBracket.BracketID
                            where tbVehicleInformation.VehicleInformationID == VehicleInformationID
                            select new VehicleInformation
                            {
                                VehicleInformationID = tbVehicleInformation.VehicleInformationID,
                                ChauffeurID = tbVehicleInformation.ChauffeurID,
                                ChauffeurName = tbChauffeur.ChauffeurName,
                                MessageID = tbVehicleInformation.MessageID,
                                ChineseName = tbMessage.ChineseName,
                                BracketID = tbVehicleInformation.BracketID,
                                BracketCode = tbBracket.BracketCode,
                                VehicleCode = tbVehicleInformation.VehicleCode,
                                PlateNumbers = tbVehicleInformation.PlateNumbers,
                                HongKongPlateNumber = tbVehicleInformation.HongKongPlateNumber,
                                VehicleClass = tbVehicleInformation.VehicleClass.Trim(),
                                MotorcycleType = tbVehicleInformation.MotorcycleType.Trim(),
                                VehicleType = tbVehicleInformation.VehicleType.Trim(),
                                Usualplace = tbVehicleInformation.Usualplace
                            }).Single();
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return null;
            }
        }
        /// <summary>
        /// 删除车辆信息
        /// </summary>
        /// <returns></returns>
        public ActionResult DelectVehicleInformation(int VehicleInformationID)
        {
            string strMsg = "fail";
            try
            {
                var listVehicleInformation = myModels.SYS_VehicleInformation.Where(m => m.VehicleInformationID == VehicleInformationID).Single();
                myModels.SYS_VehicleInformation.Remove(listVehicleInformation);
                if (myModels.SaveChanges() > 0)
                {
                    strMsg = "true";
                }
            }
            catch (Exception e)
            {
                strMsg = "fail";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 导出车辆信息
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportToExcelVehicleInformation(int ChauffeurID, int MessageID, int BracketID, string VehicleCode, string VehicleClass, string VehicleType)
        {
            var list = (from tbVehicleInformation in myModels.SYS_VehicleInformation
                        join tbChauffeur in myModels.SYS_Chauffeur on tbVehicleInformation.ChauffeurID equals tbChauffeur.ChauffeurID
                        join tbMessage in myModels.SYS_Message on tbVehicleInformation.MessageID equals tbMessage.MessageID
                        join tbBracket in myModels.SYS_Bracket on tbVehicleInformation.BracketID equals tbBracket.BracketID
                        select new VehicleInformation
                        {
                            VehicleInformationID = tbVehicleInformation.VehicleInformationID,
                            ChauffeurID = tbVehicleInformation.ChauffeurID,
                            ChauffeurName = tbChauffeur.ChauffeurName,
                            MessageID = tbVehicleInformation.MessageID,
                            ChineseName = tbMessage.ChineseName,
                            BracketID = tbVehicleInformation.BracketID,
                            BracketCode = tbBracket.BracketCode,
                            VehicleCode = tbVehicleInformation.VehicleCode,
                            PlateNumbers = tbVehicleInformation.PlateNumbers,
                            HongKongPlateNumber = tbVehicleInformation.HongKongPlateNumber,
                            VehicleClass = tbVehicleInformation.VehicleClass,
                            MotorcycleType = tbVehicleInformation.MotorcycleType,
                            VehicleType = tbVehicleInformation.VehicleType,
                            Usualplace = tbVehicleInformation.Usualplace
                        }).ToList();
            if (ChauffeurID > 0)
            {
                list = list.Where(m => m.ChauffeurID == ChauffeurID).ToList();
            }
            if (MessageID > 0)
            {
                list = list.Where(m => m.MessageID == MessageID).ToList();
            }
            if (BracketID > 0)
            {
                list = list.Where(m => m.BracketID == BracketID).ToList();
            }
            if (!string.IsNullOrEmpty(VehicleCode))
            {
                list = list.Where(m => m.VehicleCode.Contains(VehicleCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(VehicleClass))
            {
                list = list.Where(m => m.VehicleClass.Contains(VehicleClass.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(VehicleType))
            {
                list = list.Where(m => m.VehicleType.Contains(VehicleType.Trim())).ToList();
            }
            //查询数据
            List<VehicleInformation> listExaminee = list.ToList();
            //二：代码创建一个Excel表格（这里称为工作簿）
            //创建Excel文件的对象 工作簿(调用NPOI文件)
            HSSFWorkbook excelBook = new HSSFWorkbook();
            //创建Excel工作表 Sheet
            ISheet sheet1 = excelBook.CreateSheet("车辆信息");
            //给Sheet添加第一行的头部标题
            IRow row1 = sheet1.CreateRow(0);
            //给标题的每一个单元格赋值
            row1.CreateCell(0).SetCellValue("车辆代码");
            row1.CreateCell(1).SetCellValue("车牌号");
            row1.CreateCell(2).SetCellValue("司机姓名");
            row1.CreateCell(3).SetCellValue("公司信息");
            row1.CreateCell(4).SetCellValue("托架姓名");
            row1.CreateCell(5).SetCellValue("香港车牌号");
            row1.CreateCell(6).SetCellValue("车辆类别");
            row1.CreateCell(7).SetCellValue("车型");
            row1.CreateCell(8).SetCellValue("车辆类型");
            row1.CreateCell(9).SetCellValue("常驻地");
            //添加数据行：将表格数据逐步写入sheet1各个行中（也就是给每一个单元格赋值）
            for (int i = 0; i < listExaminee.Count; i++)
            {
                //创建行
                IRow rowTemp = sheet1.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listExaminee[i].VehicleCode);
                rowTemp.CreateCell(1).SetCellValue(listExaminee[i].PlateNumbers);
                rowTemp.CreateCell(2).SetCellValue(listExaminee[i].ChauffeurID);
                rowTemp.CreateCell(3).SetCellValue(listExaminee[i].MessageID);
                rowTemp.CreateCell(4).SetCellValue(listExaminee[i].BracketID);
                rowTemp.CreateCell(5).SetCellValue(listExaminee[i].HongKongPlateNumber);
                rowTemp.CreateCell(6).SetCellValue(listExaminee[i].VehicleClass);
                rowTemp.CreateCell(7).SetCellValue(listExaminee[i].MotorcycleType);
                rowTemp.CreateCell(8).SetCellValue(listExaminee[i].VehicleType);
                rowTemp.CreateCell(9).SetCellValue(listExaminee[i].Usualplace);
            }
            //输出的文件名称
            string fileName = "车辆信息" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
            //把Excel转为流，输出
            //创建文件流
            System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
            //将工作薄写入文件流
            excelBook.Write(bookStream);
            //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
            bookStream.Seek(0, System.IO.SeekOrigin.Begin);
            //Stream对象,文件类型,文件名称
            return File(bookStream, "application/vnd.ms-excel", fileName);
        }
        /// <summary>
        /// 车辆信息报表
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievementVehicleInformation(int ChauffeurID, int MessageID, int BracketID, string VehicleCode, string VehicleClass, string VehicleType)
        {
            var list = (from tbVehicleInformation in myModels.SYS_VehicleInformation
                        join tbChauffeur in myModels.SYS_Chauffeur on tbVehicleInformation.ChauffeurID equals tbChauffeur.ChauffeurID
                        join tbMessage in myModels.SYS_Message on tbVehicleInformation.MessageID equals tbMessage.MessageID
                        join tbBracket in myModels.SYS_Bracket on tbVehicleInformation.BracketID equals tbBracket.BracketID
                        select new VehicleInformation
                        {
                            VehicleInformationID = tbVehicleInformation.VehicleInformationID,
                            ChauffeurID = tbVehicleInformation.ChauffeurID,
                            ChauffeurName = tbChauffeur.ChauffeurName,
                            MessageID = tbVehicleInformation.MessageID,
                            ChineseName = tbMessage.ChineseName,
                            BracketID = tbVehicleInformation.BracketID,
                            BracketCode = tbBracket.BracketCode,
                            VehicleCode = tbVehicleInformation.VehicleCode,
                            PlateNumbers = tbVehicleInformation.PlateNumbers,
                            HongKongPlateNumber = tbVehicleInformation.HongKongPlateNumber,
                            VehicleClass = tbVehicleInformation.VehicleClass,
                            MotorcycleType = tbVehicleInformation.MotorcycleType,
                            VehicleType = tbVehicleInformation.VehicleType,
                            Usualplace = tbVehicleInformation.Usualplace
                        }).ToList();
            if (ChauffeurID > 0)
            {
                list = list.Where(m => m.ChauffeurID == ChauffeurID).ToList();
            }
            if (MessageID > 0)
            {
                list = list.Where(m => m.MessageID == MessageID).ToList();
            }
            if (BracketID > 0)
            {
                list = list.Where(m => m.BracketID == BracketID).ToList();
            }
            if (!string.IsNullOrEmpty(VehicleCode))
            {
                list = list.Where(m => m.VehicleCode.Contains(VehicleCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(VehicleClass))
            {
                list = list.Where(m => m.VehicleClass.Contains(VehicleClass.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(VehicleType))
            {
                list = list.Where(m => m.VehicleType.Contains(VehicleType.Trim())).ToList();
            }
            //把linq类型的数据listResult转化为DataTable类型数据
            DataTable dt = LINQToDataTableVehicleInformation(list);
            //第一步：实例化数据集`
            PrintReport.ReportDB dbReport = new PrintReport.ReportDB();
            //第二步：将dt的数据放入数据集的数据表中
            dbReport.Tables["VehicleInformation"].Merge(dt);
            //第三步：实例化报表模板
            PrintReport.vehicleInformation rp = new PrintReport.vehicleInformation();
            //第四步：获取报表物理文件地址
            string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                + "Areas\\Basics\\PrintReport\\vehicleInformation.rpt";
            //第五步：把报表文件加载到ReportDocument
            rp.Load(strRptPath);
            //第六步：设置报表数据源
            rp.SetDataSource(dbReport);
            //第七步：把ReportDocument转化为文件流
            Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
            return File(stream, "application/pdf");
        }
        // //将IEnumerable<T>类型的集合转换为DataTable类型
        public DataTable LINQToDataTableVehicleInformation<T>(IEnumerable<T> varlist)
        {
            //定义要返回的DataTable对象
            DataTable dtReturn = new DataTable();
            //保存列集合的属性信息数组
            PropertyInfo[] oProps = null;
            if (varlist == null) return dtReturn;//安全性检查
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
                        //得到属性的类型
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
        #region 托架资料
        /// <summary>
        /// 查询托架资料
        /// </summary>
        /// <returns></returns>
        public ActionResult selectBrackett(BsgridPage bsgridPage, short? MessageID, string BracketCode, string BracketTag)
        {
            var list = (from tbBracket in myModels.SYS_Bracket
                        join tbMessage in myModels.SYS_Message on tbBracket.MessageID equals tbMessage.MessageID
                        select new Bracket
                        {
                            MessageID = tbBracket.MessageID,
                            ChineseName = tbMessage.ChineseName,
                            BracketID = tbBracket.BracketID,
                            BracketCode = tbBracket.BracketCode,
                            ChassisNumber = tbBracket.ChassisNumber,
                            BracketType = tbBracket.BracketType,
                            CustomsRecordModel = tbBracket.CustomsRecordModel,
                            BracketTag = tbBracket.BracketTag,
                            SinceNumber = tbBracket.SinceNumber,
                            timer = tbBracket.ManufactureDate.ToString(),
                            ManufactureDate =tbBracket.ManufactureDate,
                            ManufactureFactory=tbBracket.ManufactureFactory,
                            WhetherStart=tbBracket.WhetherStart
                        }).ToList();
            if (MessageID > 0)
            {
                list = list.Where(m => m.MessageID == MessageID).ToList();
            }
            if (!string.IsNullOrEmpty(BracketCode))
            {
                list = list.Where(m => m.BracketCode.Contains(BracketCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(BracketTag))
            {
                list = list.Where(m => m.BracketTag.Contains(BracketTag.Trim())).ToList();
            }
            int totalRow = list.Count();
            List<Bracket> notices = list.OrderByDescending(p => p.BracketID).
              Skip(bsgridPage.GetStartIndex()).
              Take(bsgridPage.pageSize).
              ToList();
            Bsgrid<Bracket> bsgrid = new Bsgrid<Bracket>();
            bsgrid.success = true;
            bsgrid.totalRows = totalRow;
            bsgrid.curPage = bsgridPage.curPage;
            bsgrid.data = notices;
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 新增托架资料
        /// </summary>
        /// <returns></returns>
        public ActionResult InsertBracket(SYS_Bracket tbbracket)
        {
            string strMsg = "fail";
            try
            {
                int oldBracket = (from tbBracket in myModels.SYS_Bracket
                                  where tbBracket.BracketCode == tbbracket.BracketCode
                                  select tbBracket).Count();
                if (oldBracket == 0)
                {
                    int BracketTag = (from tbBracket in myModels.SYS_Bracket
                                      where tbBracket.BracketTag == tbbracket.BracketTag
                                      select tbBracket).Count();
                    if (BracketTag == 0)
                    {
                        myModels.SYS_Bracket.Add(tbbracket);
                        if (myModels.SaveChanges() > 0)
                        {
                            strMsg = "数据新增成功！";
                        }
                    }
                    else
                    {
                        strMsg = "已有托架尾牌号，不能新增！";
                    }
                }
                else
                {
                    strMsg = "已有托架代码，不能新增！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 修改托架资料
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateBracket(SYS_Bracket pwBracket)
        {
            string strMsg = "fail";
            try
            {
                if (myModels.SYS_Bracket.Where(m => m.BracketCode == pwBracket.BracketCode && m.BracketID != pwBracket.BracketID).Count() == 0)
                {
                    if (myModels.SYS_Bracket.Where(m => m.BracketTag == pwBracket.BracketTag && m.BracketID != pwBracket.BracketID).Count() == 0)
                    {
                        SYS_Bracket dbBracket = (from tbBracket in myModels.SYS_Bracket
                                                 where tbBracket.BracketID == pwBracket.BracketID
                                                 select tbBracket).Single();
                        dbBracket.MessageID = pwBracket.MessageID;
                        dbBracket.BracketCode = pwBracket.BracketCode;
                        dbBracket.ChassisNumber = pwBracket.ChassisNumber;
                        dbBracket.BracketType = pwBracket.BracketType;
                        dbBracket.CustomsRecordModel = pwBracket.CustomsRecordModel;
                        dbBracket.BracketTag = pwBracket.BracketTag;
                        dbBracket.SinceNumber = pwBracket.SinceNumber;
                        dbBracket.ManufactureDate = pwBracket.ManufactureDate;
                        dbBracket.ManufactureFactory = pwBracket.ManufactureFactory;
                        dbBracket.WhetherStart = pwBracket.WhetherStart;
                        myModels.Entry(dbBracket).State = System.Data.Entity.EntityState.Modified;
                        if (myModels.SaveChanges() > 0)
                        {
                            strMsg = "修改成功！";
                        }
                    }
                    else
                    {
                        strMsg = "已有托架尾牌号，无法修改！";
                    }
                }
                else
                {
                    strMsg = "已有托架代码，无法修改！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 回填托架资料
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectBracketById(int BracketID)
        {
            try
            {
                var list = (from tbBracket in myModels.SYS_Bracket
                            join tbMessage in myModels.SYS_Message on tbBracket.MessageID equals tbMessage.MessageID
                            where tbBracket.BracketID == BracketID
                            select new Bracket
                            {
                                MessageID = tbBracket.MessageID,
                                ChineseName = tbMessage.ChineseName,
                                BracketID = tbBracket.BracketID,
                                BracketCode = tbBracket.BracketCode,
                                ChassisNumber = tbBracket.ChassisNumber,
                                BracketType = tbBracket.BracketType,
                                CustomsRecordModel = tbBracket.CustomsRecordModel,
                                BracketTag = tbBracket.BracketTag,
                                SinceNumber = tbBracket.SinceNumber,
                                timer = tbBracket.ManufactureDate.ToString(),
                                ManufactureDate = tbBracket.ManufactureDate,
                                ManufactureFactory = tbBracket.ManufactureFactory,
                                WhetherStart = tbBracket.WhetherStart
                            }).Single();
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return null;
            }
        }
        /// <summary>
        /// 删除托架资料
        /// </summary>
        /// <returns></returns>
        public ActionResult DelectBracket(int BracketID)
        {
            string strMsg = "fail";
            try
            {
                var listBracket = myModels.SYS_Bracket.Where(m => m.BracketID == BracketID).Single();
                myModels.SYS_Bracket.Remove(listBracket);
                if (myModels.SaveChanges() > 0)
                {
                    strMsg = "true";
                }
            }
            catch (Exception e)
            {
                strMsg = "fail";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 导出托架资料
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportToExcelBracket(short? MessageID, string BracketCode, string BracketTag)
        {
            var list = (from tbBracket in myModels.SYS_Bracket
                        join tbMessage in myModels.SYS_Message on tbBracket.MessageID equals tbMessage.MessageID
                        select new Bracket
                        {
                            MessageID = tbBracket.MessageID,
                            ChineseName = tbMessage.ChineseName,
                            BracketID = tbBracket.BracketID,
                            BracketCode = tbBracket.BracketCode,
                            ChassisNumber = tbBracket.ChassisNumber,
                            BracketType = tbBracket.BracketType,
                            CustomsRecordModel = tbBracket.CustomsRecordModel,
                            BracketTag = tbBracket.BracketTag,
                            SinceNumber = tbBracket.SinceNumber,
                            timer = tbBracket.ManufactureDate.ToString(),
                            ManufactureDate = tbBracket.ManufactureDate,
                            ManufactureFactory = tbBracket.ManufactureFactory,
                            WhetherStartb = tbBracket.WhetherStart,
                            WhetherStarta = ""
                        }).ToList();
            if (MessageID > 0)
            {
                list = list.Where(m => m.MessageID == MessageID).ToList();
            }
            if (!string.IsNullOrEmpty(BracketCode))
            {
                list = list.Where(m => m.BracketCode.Contains(BracketCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(BracketTag))
            {
                list = list.Where(m => m.BracketTag.Contains(BracketTag.Trim())).ToList();
            }
            for (int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherStartb == true)
                {
                    list[i].WhetherStarta = "启用";
                }
                else if (list[i].WhetherStartb == false)
                {
                    list[i].WhetherStarta = "不启用";
                }
            }
            //查询数据
            List<Bracket> listExaminee = list.ToList();
            //二：代码创建一个Excel表格（这里称为工作簿）
            //创建Excel文件的对象 工作簿(调用NPOI文件)
            HSSFWorkbook excelBook = new HSSFWorkbook();
            //创建Excel工作表 Sheet
            ISheet sheet1 = excelBook.CreateSheet("托架资料");
            //给Sheet添加第一行的头部标题
            IRow row1 = sheet1.CreateRow(0);
            //给标题的每一个单元格赋值
            row1.CreateCell(0).SetCellValue("托架代码");
            row1.CreateCell(1).SetCellValue("底盘号");
            row1.CreateCell(2).SetCellValue("托架型号");
            row1.CreateCell(3).SetCellValue("所属机构");
            row1.CreateCell(4).SetCellValue("海关备案型号");
            row1.CreateCell(5).SetCellValue("托架尾牌号");
            row1.CreateCell(6).SetCellValue("自编号");
            row1.CreateCell(7).SetCellValue("生产日期");
            row1.CreateCell(8).SetCellValue("生产厂家");
            row1.CreateCell(9).SetCellValue("是否启用");
            //添加数据行：将表格数据逐步写入sheet1各个行中（也就是给每一个单元格赋值）
            for (int i = 0; i < listExaminee.Count; i++)
            {
                //创建行
                IRow rowTemp = sheet1.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listExaminee[i].BracketCode);
                rowTemp.CreateCell(1).SetCellValue(listExaminee[i].ChassisNumber);
                rowTemp.CreateCell(2).SetCellValue(listExaminee[i].BracketType);
                rowTemp.CreateCell(3).SetCellValue(listExaminee[i].ChineseName);
                rowTemp.CreateCell(4).SetCellValue(listExaminee[i].CustomsRecordModel);
                rowTemp.CreateCell(5).SetCellValue(listExaminee[i].BracketTag);
                rowTemp.CreateCell(6).SetCellValue(listExaminee[i].SinceNumber);
                rowTemp.CreateCell(7).SetCellValue(listExaminee[i].ManufactureDate.ToString());
                rowTemp.CreateCell(8).SetCellValue(listExaminee[i].ManufactureFactory);
                rowTemp.CreateCell(9).SetCellValue(listExaminee[i].WhetherStarta);
            }
            //输出的文件名称
            string fileName = "托架资料" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
            //把Excel转为流，输出
            //创建文件流
            System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
            //将工作薄写入文件流
            excelBook.Write(bookStream);
            //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
            bookStream.Seek(0, System.IO.SeekOrigin.Begin);
            //Stream对象,文件类型,文件名称
            return File(bookStream, "application/vnd.ms-excel", fileName);
        }
        /// <summary>
        /// 托架资料报表
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievementBracket(short? MessageID, string BracketCode, string BracketTag)
        {
            var list = (from tbBracket in myModels.SYS_Bracket
                        join tbMessage in myModels.SYS_Message on tbBracket.MessageID equals tbMessage.MessageID
                        select new Bracket
                        {
                            MessageID = tbBracket.MessageID,
                            ChineseName = tbMessage.ChineseName,
                            BracketID = tbBracket.BracketID,
                            BracketCode = tbBracket.BracketCode,
                            ChassisNumber = tbBracket.ChassisNumber,
                            BracketType = tbBracket.BracketType,
                            CustomsRecordModel = tbBracket.CustomsRecordModel,
                            BracketTag = tbBracket.BracketTag,
                            SinceNumber = tbBracket.SinceNumber,
                            timer = tbBracket.ManufactureDate.ToString(),
                            ManufactureDate = tbBracket.ManufactureDate,
                            ManufactureFactory = tbBracket.ManufactureFactory,
                            WhetherStartb = tbBracket.WhetherStart,
                            WhetherStarta = ""
                        }).ToList();
            if (MessageID > 0)
            {
                list = list.Where(m => m.MessageID == MessageID).ToList();
            }
            if (!string.IsNullOrEmpty(BracketCode))
            {
                list = list.Where(m => m.BracketCode.Contains(BracketCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(BracketTag))
            {
                list = list.Where(m => m.BracketTag.Contains(BracketTag.Trim())).ToList();
            }
            for (int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherStartb == true)
                {
                    list[i].WhetherStarta = "√";
                }
                else if (list[i].WhetherStartb == false)
                {
                    list[i].WhetherStarta = "×";
                }
            }
            //把linq类型的数据listResult转化为DataTable类型数据
            DataTable dt = LINQToDataTableBracket(list);
            //第一步：实例化数据集`
            PrintReport.ReportDB dbReport = new PrintReport.ReportDB();
            //第二步：将dt的数据放入数据集的数据表中
            dbReport.Tables["Bracket"].Merge(dt);
            //第三步：实例化报表模板
            PrintReport.bracket rp = new PrintReport.bracket();
            //第四步：获取报表物理文件地址
            string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                + "Areas\\Basics\\PrintReport\\bracket.rpt";
            //第五步：把报表文件加载到ReportDocument
            rp.Load(strRptPath);
            //第六步：设置报表数据源
            rp.SetDataSource(dbReport);
            //第七步：把ReportDocument转化为文件流
            Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
            return File(stream, "application/pdf");
        }
        // //将IEnumerable<T>类型的集合转换为DataTable类型
        public DataTable LINQToDataTableBracket<T>(IEnumerable<T> varlist)
        {
            //定义要返回的DataTable对象
            DataTable dtReturn = new DataTable();
            //保存列集合的属性信息数组
            PropertyInfo[] oProps = null;
            if (varlist == null) return dtReturn;//安全性检查
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
                        //得到属性的类型
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
        #region 提还柜地
        /// <summary>
        /// 查询提还柜地
        /// </summary>
        /// <returns></returns>
        public ActionResult selectMention(BsgridPage bsgridPage, int? GatedotID, string MCode, string Abbreviation)
        {
            var list = (from tbMention in myModels.SYS_Mention
                        join tbGatedot in myModels.SYS_Gatedot on tbMention.GatedotID equals tbGatedot.GatedotID
                        select new Mention
                        {
                            GatedotID= tbMention.GatedotID,
                            GatedotName=tbGatedot.GatedotName,
                            MentionID = tbMention.MentionID,
                            MCode=tbMention.MCode,
                            Abbreviation=tbMention.Abbreviation,
                            Linkman=tbMention.Linkman,
                            Mobile=tbMention.Mobile,
                            Site=tbMention.Site,
                            WhetherStart=tbMention.WhetherStart,
                            HangCharge=tbMention.HangCharge,
                            CheckFee=tbMention.CheckFee,
                            GateCharges=tbMention.GateCharges
                        }).ToList();
            if (GatedotID > 0)
            {
                list = list.Where(m => m.GatedotID == GatedotID).ToList();
            }
            if (!string.IsNullOrEmpty(MCode))
            {
                list = list.Where(m => m.MCode.Contains(MCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(Abbreviation))
            {
                list = list.Where(m => m.Abbreviation.Contains(Abbreviation.Trim())).ToList();
            }
            int totalRow = list.Count();
            List<Mention> notices = list.OrderByDescending(p => p.MentionID).
              Skip(bsgridPage.GetStartIndex()).
              Take(bsgridPage.pageSize).
              ToList();
            Bsgrid<Mention> bsgrid = new Bsgrid<Mention>();
            bsgrid.success = true;
            bsgrid.totalRows = totalRow;
            bsgrid.curPage = bsgridPage.curPage;
            bsgrid.data = notices;
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 新增提还柜地
        /// </summary>
        /// <returns></returns>
        public ActionResult InsertMention(SYS_Mention tbmention)
        {
            string strMsg = "fail";
            try
            {
                int MCode = (from tbMention in myModels.SYS_Mention
                             where tbMention.MCode == tbmention.MCode
                             select tbMention).Count();
                if (MCode == 0)
                {
                    int Abbreviation = (from tbMention in myModels.SYS_Mention
                                        where tbMention.Abbreviation == tbmention.Abbreviation
                                        select tbMention).Count();
                    if (Abbreviation == 0)
                    {
                        myModels.SYS_Mention.Add(tbmention);
                        if (myModels.SaveChanges() > 0)
                        {
                            strMsg = "数据新增成功！";
                        }
                    }
                    else
                    {
                        strMsg = "已有简称，不能新增！";
                    }
                }
                else
                {
                    strMsg = "已有编码，不能新增！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 修改提还柜地
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateMention(SYS_Mention pwMention)
        {
            string strMsg = "fail";
            try
            {
                if (myModels.SYS_Mention.Where(m => m.MCode == pwMention.MCode && m.MentionID != pwMention.MentionID).Count() == 0)
                {
                    if (myModels.SYS_Mention.Where(m => m.Abbreviation == pwMention.Abbreviation && m.MentionID != pwMention.MentionID).Count() == 0)
                    {
                        SYS_Mention dbMention = (from tbMention in myModels.SYS_Mention
                                                 where tbMention.MentionID == pwMention.MentionID
                                                 select tbMention).Single();
                        dbMention.GatedotID = pwMention.GatedotID;
                        dbMention.Abbreviation = pwMention.Abbreviation;
                        dbMention.CheckFee = pwMention.CheckFee;
                        dbMention.MCode = pwMention.MCode;
                        dbMention.Linkman = pwMention.Linkman;
                        dbMention.Mobile = pwMention.Mobile;
                        dbMention.Site = pwMention.Site;
                        dbMention.WhetherStart = pwMention.WhetherStart;
                        dbMention.HangCharge = pwMention.HangCharge;
                        dbMention.GateCharges = pwMention.GateCharges;
                        myModels.Entry(dbMention).State = System.Data.Entity.EntityState.Modified;
                        if (myModels.SaveChanges() > 0)
                        {
                            strMsg = "修改成功！";
                        }
                    }
                    else
                    {
                        strMsg = "已有简称，无法修改！";
                    }
                }
                else
                {
                    strMsg = "已有编码，无法修改！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 回填提还柜地
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectMentionById(int MentionID)
        {
            try
            {
                var list = (from tbMention in myModels.SYS_Mention
                            join tbGatedot in myModels.SYS_Gatedot on tbMention.GatedotID equals tbGatedot.GatedotID
                            where tbMention.MentionID== MentionID
                            select new Mention
                            {
                                GatedotID = tbMention.GatedotID,
                                GatedotName = tbGatedot.GatedotName,
                                MentionID = tbMention.MentionID,
                                MCode = tbMention.MCode,
                                Abbreviation = tbMention.Abbreviation,
                                Linkman = tbMention.Linkman,
                                Mobile = tbMention.Mobile,
                                Site = tbMention.Site,
                                WhetherStart = tbMention.WhetherStart,
                                HangCharge = tbMention.HangCharge,
                                CheckFee = tbMention.CheckFee,
                                GateCharges = tbMention.GateCharges
                            }).Single();
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return null;
            }
        }
        /// <summary>
        /// 删除提还柜地
        /// </summary>
        /// <returns></returns>
        public ActionResult DelectMention(int MentionID)
        {
            string strMsg = "fail";
            try
            {
                var listMention = myModels.SYS_Mention.Where(m => m.MentionID == MentionID).Single();
                myModels.SYS_Mention.Remove(listMention);
                if (myModels.SaveChanges() > 0)
                {
                    strMsg = "true";
                }
            }
            catch (Exception e)
            {
                strMsg = "fail";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 导出提还柜地
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportToExcelMention(int GatedotID, string MCode, string Abbreviation)
        {
            var list = (from tbMention in myModels.SYS_Mention
                        join tbGatedot in myModels.SYS_Gatedot on tbMention.GatedotID equals tbGatedot.GatedotID
                        select new Mention
                        {
                            GatedotID = tbMention.GatedotID,
                            GatedotName = tbGatedot.GatedotName,
                            MentionID = tbMention.MentionID,
                            MCode = tbMention.MCode,
                            Abbreviation = tbMention.Abbreviation,
                            Linkman = tbMention.Linkman,
                            Mobile = tbMention.Mobile,
                            Site = tbMention.Site,
                            WhetherStarta="",
                            WhetherStartb = tbMention.WhetherStart,
                            HangCharge = tbMention.HangCharge,
                            CheckFee = tbMention.CheckFee,
                            GateCharges = tbMention.GateCharges
                        }).ToList();
            if (GatedotID > 0)
            {
                list = list.Where(m => m.GatedotID == GatedotID).ToList();
            }
            if (!string.IsNullOrEmpty(MCode))
            {
                list = list.Where(m => m.MCode.Contains(MCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(Abbreviation))
            {
                list = list.Where(m => m.Abbreviation.Contains(Abbreviation.Trim())).ToList();
            }
            for (int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherStartb == true)
                {
                    list[i].WhetherStarta = "启用";
                }
                else if (list[i].WhetherStartb == false)
                {
                    list[i].WhetherStarta = "不启用";
                }
            }
            //查询数据
            List<Mention> listExaminee = list.ToList();
            //二：代码创建一个Excel表格（这里称为工作簿）
            //创建Excel文件的对象 工作簿(调用NPOI文件)
            HSSFWorkbook excelBook = new HSSFWorkbook();
            //创建Excel工作表 Sheet
            ISheet sheet1 = excelBook.CreateSheet("提还柜地");
            //给Sheet添加第一行的头部标题
            IRow row1 = sheet1.CreateRow(0);
            //给标题的每一个单元格赋值
            row1.CreateCell(0).SetCellValue("编码");
            row1.CreateCell(1).SetCellValue("简称");
            row1.CreateCell(2).SetCellValue("所属区域");
            row1.CreateCell(3).SetCellValue("联系人");
            row1.CreateCell(4).SetCellValue("电话");
            row1.CreateCell(5).SetCellValue("地址");
            row1.CreateCell(6).SetCellValue("吊柜费");
            row1.CreateCell(7).SetCellValue("查柜费");
            row1.CreateCell(8).SetCellValue("闸口费");
            row1.CreateCell(9).SetCellValue("是否启用");
            //添加数据行：将表格数据逐步写入sheet1各个行中（也就是给每一个单元格赋值）
            for (int i = 0; i < listExaminee.Count; i++)
            {
                //创建行
                IRow rowTemp = sheet1.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listExaminee[i].MCode);
                rowTemp.CreateCell(1).SetCellValue(listExaminee[i].Abbreviation);
                rowTemp.CreateCell(2).SetCellValue(listExaminee[i].GatedotName);
                rowTemp.CreateCell(3).SetCellValue(listExaminee[i].Linkman);
                rowTemp.CreateCell(4).SetCellValue(listExaminee[i].Mobile);
                rowTemp.CreateCell(5).SetCellValue(listExaminee[i].Site);
                rowTemp.CreateCell(6).SetCellValue(listExaminee[i].HangCharge.ToString());
                rowTemp.CreateCell(7).SetCellValue(listExaminee[i].CheckFee.ToString());
                rowTemp.CreateCell(8).SetCellValue(listExaminee[i].GateCharges.ToString());
                rowTemp.CreateCell(9).SetCellValue(listExaminee[i].WhetherStarta);
            }
            //输出的文件名称
            string fileName = "提还柜地" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
            //把Excel转为流，输出
            //创建文件流
            System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
            //将工作薄写入文件流
            excelBook.Write(bookStream);
            //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
            bookStream.Seek(0, System.IO.SeekOrigin.Begin);
            //Stream对象,文件类型,文件名称
            return File(bookStream, "application/vnd.ms-excel", fileName);
        }
        /// <summary>
        /// 提还柜地报表
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievementMention(int GatedotID, string MCode, string Abbreviation)
        {
            var list = (from tbMention in myModels.SYS_Mention
                        join tbGatedot in myModels.SYS_Gatedot on tbMention.GatedotID equals tbGatedot.GatedotID
                        select new Mention
                        {
                            GatedotID = tbMention.GatedotID,
                            GatedotName = tbGatedot.GatedotName,
                            MentionID = tbMention.MentionID,
                            MCode = tbMention.MCode,
                            Abbreviation = tbMention.Abbreviation,
                            Linkman = tbMention.Linkman,
                            Mobile = tbMention.Mobile,
                            Site = tbMention.Site,
                            WhetherStarta="",
                            WhetherStartb = tbMention.WhetherStart,
                            HangCharge = tbMention.HangCharge,
                            CheckFee = tbMention.CheckFee,
                            GateCharges = tbMention.GateCharges
                        }).ToList();
            if (GatedotID > 0)
            {
                list = list.Where(m => m.GatedotID == GatedotID).ToList();
            }
            if (!string.IsNullOrEmpty(MCode))
            {
                list = list.Where(m => m.MCode.Contains(MCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(Abbreviation))
            {
                list = list.Where(m => m.Abbreviation.Contains(Abbreviation.Trim())).ToList();
            }
            for (int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherStartb == true)
                {
                    list[i].WhetherStarta = "√";
                }
                else if (list[i].WhetherStartb == false)
                {
                    list[i].WhetherStarta = "×";
                }
            }
            //把linq类型的数据listResult转化为DataTable类型数据
            DataTable dt = LINQToDataTableMention(list);
            //第一步：实例化数据集`
            PrintReport.ReportDB dbReport = new PrintReport.ReportDB();
            //第二步：将dt的数据放入数据集的数据表中
            dbReport.Tables["Mention"].Merge(dt);
            //第三步：实例化报表模板
            PrintReport.mention rp = new PrintReport.mention();
            //第四步：获取报表物理文件地址
            string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                + "Areas\\Basics\\PrintReport\\mention.rpt";
            //第五步：把报表文件加载到ReportDocument
            rp.Load(strRptPath);
            //第六步：设置报表数据源
            rp.SetDataSource(dbReport);
            //第七步：把ReportDocument转化为文件流
            Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
            return File(stream, "application/pdf");
        }
        // //将IEnumerable<T>类型的集合转换为DataTable类型
        public DataTable LINQToDataTableMention<T>(IEnumerable<T> varlist)
        {
            //定义要返回的DataTable对象
            DataTable dtReturn = new DataTable();
            //保存列集合的属性信息数组
            PropertyInfo[] oProps = null;
            if (varlist == null) return dtReturn;//安全性检查
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
                        //得到属性的类型
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
        #region 关区管理
        /// <summary>
        /// 查询关区管理
        /// </summary>
        /// <returns></returns>
        public ActionResult selectCustoms(BsgridPage bsgridPage,int? GatedotID,string CustomsName,string CustomsCode)
        {
            var list = (from tbCustoms in myModels.SYS_Customs
                        join tbGatedot in myModels.SYS_Gatedot on tbCustoms.GatedotID equals tbGatedot.GatedotID
                        select new Customs
                        {
                            CustomsID = tbCustoms.CustomsID,
                            GatedotID = tbCustoms.GatedotID,
                            GatedotName = tbGatedot.GatedotName,
                            CustomsName = tbCustoms.CustomsName,
                            CustomsCode = tbCustoms.CustomsCode,
                            EnglishName = tbCustoms.EnglishName,
                            MobilePhone = tbCustoms.MobilePhone,
                            Mobile = tbCustoms.Mobile,
                            Linkman = tbCustoms.Linkman,
                            Email = tbCustoms.Email,
                            Faxes = tbCustoms.Faxes,
                            PostCode = tbCustoms.PostCode,
                            OfficeHours = tbCustoms.OfficeHours,
                            timer= tbCustoms.OfficeHours.ToString(),
                            ClosingTime = tbCustoms.ClosingTime,
                            timerer= tbCustoms.ClosingTime.ToString(),
                            Site = tbCustoms.Site,
                            Website = tbCustoms.Website,
                            Describe = tbCustoms.Describe,
                            WhetherStart = tbCustoms.WhetherStart
                        }).ToList();
            if (GatedotID > 0)
            {
                list = list.Where(m => m.GatedotID == GatedotID).ToList();
            }
            if (!string.IsNullOrEmpty(CustomsName))
            {
                list = list.Where(m => m.CustomsName.Contains(CustomsName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(CustomsCode))
            {
                list = list.Where(m => m.CustomsCode.Contains(CustomsCode.Trim())).ToList();
            }
            int totalRow = list.Count();
            List<Customs> notices = list.OrderByDescending(p => p.CustomsID).
              Skip(bsgridPage.GetStartIndex()).
              Take(bsgridPage.pageSize).
              ToList();
            Bsgrid<Customs> bsgrid = new Bsgrid<Customs>();
            bsgrid.success = true;
            bsgrid.totalRows = totalRow;
            bsgrid.curPage = bsgridPage.curPage;
            bsgrid.data = notices;
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 新增关区管理
        /// </summary>
        /// <returns></returns>
        public ActionResult InsertCustoms(SYS_Customs tbcustoms)
        {
            string strMsg = "fail";
            try
            {
                int Code = (from tbCustoms in myModels.SYS_Customs
                                  where tbCustoms.CustomsCode == tbcustoms.CustomsCode
                                  select tbCustoms).Count();
                if (Code == 0)
                {
                    int Name = (from tbCustoms in myModels.SYS_Customs
                                      where tbCustoms.CustomsName == tbcustoms.CustomsName
                                      select tbCustoms).Count();
                    if (Name == 0)
                    {
                        myModels.SYS_Customs.Add(tbcustoms);
                        if (myModels.SaveChanges() > 0)
                        {
                            strMsg = "数据新增成功！";
                        }
                    }
                    else
                    {
                        strMsg = "已有关区名称，不能新增！";
                    }
                }
                else
                {
                    strMsg = "已有关区编码，不能新增！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 修改关区管理
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateCustoms(SYS_Customs pwcustoms)
        {
            string strMsg = "fail";
            try
            {
                if (myModels.SYS_Customs.Where(m => m.CustomsCode == pwcustoms.CustomsCode && m.CustomsID != pwcustoms.CustomsID).Count() == 0)
                {
                    if (myModels.SYS_Customs.Where(m => m.CustomsName == pwcustoms.CustomsName && m.CustomsID != pwcustoms.CustomsID).Count() == 0)
                    {
                        SYS_Customs dbCustoms = (from tbCustoms in myModels.SYS_Customs
                                                 where tbCustoms.CustomsID == pwcustoms.CustomsID
                                                 select tbCustoms).Single();
                        dbCustoms.GatedotID = pwcustoms.GatedotID;
                        dbCustoms.CustomsName = pwcustoms.CustomsName;
                        dbCustoms.CustomsCode = pwcustoms.CustomsCode;
                        dbCustoms.EnglishName = pwcustoms.EnglishName;
                        dbCustoms.MobilePhone = pwcustoms.MobilePhone;
                        dbCustoms.Mobile = pwcustoms.Mobile;
                        dbCustoms.Linkman = pwcustoms.Linkman;
                        dbCustoms.Email = pwcustoms.Email;
                        dbCustoms.Faxes = pwcustoms.Faxes;
                        dbCustoms.PostCode = pwcustoms.PostCode;
                        dbCustoms.OfficeHours = pwcustoms.OfficeHours;
                        dbCustoms.ClosingTime = pwcustoms.ClosingTime;
                        dbCustoms.Site = pwcustoms.Site;
                        dbCustoms.Website = pwcustoms.Website;
                        dbCustoms.WhetherStart = pwcustoms.WhetherStart;
                        dbCustoms.Describe = pwcustoms.Describe;
                        myModels.Entry(dbCustoms).State = System.Data.Entity.EntityState.Modified;
                        if (myModels.SaveChanges() > 0)
                        {
                            strMsg = "修改成功！";
                        }
                    }
                    else
                    {
                        strMsg = "已有关区名称，无法修改！";
                    }
                }
                else
                {
                    strMsg = "已有关区编码，无法修改！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 回填关区管理
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectCustomsById(int CustomsID)
        {
            try
            {
                var list = (from tbCustoms in myModels.SYS_Customs
                            join tbGatedot in myModels.SYS_Gatedot on tbCustoms.GatedotID equals tbGatedot.GatedotID
                            where tbCustoms.CustomsID== CustomsID
                            select new Customs
                            {
                                CustomsID = tbCustoms.CustomsID,
                                GatedotID = tbCustoms.GatedotID,
                                GatedotName = tbGatedot.GatedotName,
                                CustomsName = tbCustoms.CustomsName,
                                CustomsCode = tbCustoms.CustomsCode,
                                EnglishName = tbCustoms.EnglishName,
                                MobilePhone = tbCustoms.MobilePhone,
                                Mobile = tbCustoms.Mobile,
                                Linkman = tbCustoms.Linkman,
                                Email = tbCustoms.Email,
                                Faxes = tbCustoms.Faxes,
                                PostCode = tbCustoms.PostCode,
                                OfficeHours = tbCustoms.OfficeHours,
                                ClosingTime = tbCustoms.ClosingTime,
                                Site = tbCustoms.Site,
                                Website = tbCustoms.Website,
                                Describe = tbCustoms.Describe,
                                WhetherStart = tbCustoms.WhetherStart
                            }).Single();
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return null;
            }
        }
        /// <summary>
        /// 删除关区管理
        /// </summary>
        /// <returns></returns>
        public ActionResult DelectCustoms(int CustomsID)
        {
            string strMsg = "fail";
            try
            {
                var listCustoms = myModels.SYS_Customs.Where(m => m.CustomsID == CustomsID).Single();
                myModels.SYS_Customs.Remove(listCustoms);
                if (myModels.SaveChanges() > 0)
                {
                    strMsg = "true";
                }
            }
            catch (Exception e)
            {
                strMsg = "fail";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 导出关区管理
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportToExcelCustoms(int GatedotID, string CustomsName, string CustomsCode)
        {
            var list = (from tbCustoms in myModels.SYS_Customs
                        join tbGatedot in myModels.SYS_Gatedot on tbCustoms.GatedotID equals tbGatedot.GatedotID
                        select new Customs
                        {
                            CustomsID = tbCustoms.CustomsID,
                            GatedotID = tbCustoms.GatedotID,
                            GatedotName = tbGatedot.GatedotName,
                            CustomsName = tbCustoms.CustomsName,
                            CustomsCode = tbCustoms.CustomsCode,
                            EnglishName = tbCustoms.EnglishName,
                            MobilePhone = tbCustoms.MobilePhone,
                            Mobile = tbCustoms.Mobile,
                            Linkman = tbCustoms.Linkman,
                            Email = tbCustoms.Email,
                            Faxes = tbCustoms.Faxes,
                            PostCode = tbCustoms.PostCode,
                            OfficeHours = tbCustoms.OfficeHours,
                            ClosingTime = tbCustoms.ClosingTime,
                            Site = tbCustoms.Site,
                            Website = tbCustoms.Website,
                            Describe = tbCustoms.Describe,
                            WhetherStarta="",
                            WhetherStartb = tbCustoms.WhetherStart
                        }).ToList();
            if (GatedotID > 0)
            {
                list = list.Where(m => m.GatedotID == GatedotID).ToList();
            }
            if (!string.IsNullOrEmpty(CustomsName))
            {
                list = list.Where(m => m.CustomsName.Contains(CustomsName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(CustomsCode))
            {
                list = list.Where(m => m.CustomsCode.Contains(CustomsCode.Trim())).ToList();
            }
            for (int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherStartb == true)
                {
                    list[i].WhetherStarta = "启用";
                }
                else if (list[i].WhetherStartb == false)
                {
                    list[i].WhetherStarta = "不启用";
                }
            }
            //查询数据
            List<Customs> listExaminee = list.ToList();
            //二：代码创建一个Excel表格（这里称为工作簿）
            //创建Excel文件的对象 工作簿(调用NPOI文件)
            HSSFWorkbook excelBook = new HSSFWorkbook();
            //创建Excel工作表 Sheet
            ISheet sheet1 = excelBook.CreateSheet("关区管理");
            //给Sheet添加第一行的头部标题
            IRow row1 = sheet1.CreateRow(0);
            //给标题的每一个单元格赋值
            row1.CreateCell(0).SetCellValue("关区名称");
            row1.CreateCell(1).SetCellValue("关区编码");
            row1.CreateCell(2).SetCellValue("所属区域");
            row1.CreateCell(3).SetCellValue("英文名称");
            row1.CreateCell(4).SetCellValue("手机");
            row1.CreateCell(5).SetCellValue("电话");
            row1.CreateCell(6).SetCellValue("联系人");
            row1.CreateCell(7).SetCellValue("EMAIL");
            row1.CreateCell(8).SetCellValue("传真");
            row1.CreateCell(9).SetCellValue("邮编");
            row1.CreateCell(10).SetCellValue("上班时间");
            row1.CreateCell(11).SetCellValue("下班时间");
            row1.CreateCell(12).SetCellValue("地址");
            row1.CreateCell(13).SetCellValue("网站");
            row1.CreateCell(14).SetCellValue("是否启用");
            row1.CreateCell(15).SetCellValue("描述");
            //添加数据行：将表格数据逐步写入sheet1各个行中（也就是给每一个单元格赋值）
            for (int i = 0; i < listExaminee.Count; i++)
            {
                //创建行
                IRow rowTemp = sheet1.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listExaminee[i].CustomsName);
                rowTemp.CreateCell(1).SetCellValue(listExaminee[i].CustomsCode);
                rowTemp.CreateCell(2).SetCellValue(listExaminee[i].GatedotName);
                rowTemp.CreateCell(3).SetCellValue(listExaminee[i].EnglishName);
                rowTemp.CreateCell(4).SetCellValue(listExaminee[i].MobilePhone);
                rowTemp.CreateCell(5).SetCellValue(listExaminee[i].Mobile);
                rowTemp.CreateCell(6).SetCellValue(listExaminee[i].Linkman);
                rowTemp.CreateCell(7).SetCellValue(listExaminee[i].Email);
                rowTemp.CreateCell(8).SetCellValue(listExaminee[i].Faxes);
                rowTemp.CreateCell(9).SetCellValue(listExaminee[i].PostCode);
                rowTemp.CreateCell(10).SetCellValue(listExaminee[i].OfficeHours.ToString());
                rowTemp.CreateCell(11).SetCellValue(listExaminee[i].ClosingTime.ToString());
                rowTemp.CreateCell(12).SetCellValue(listExaminee[i].Site);
                rowTemp.CreateCell(13).SetCellValue(listExaminee[i].Website);
                rowTemp.CreateCell(14).SetCellValue(listExaminee[i].WhetherStarta);
                rowTemp.CreateCell(15).SetCellValue(listExaminee[i].Describe);
            }
            //输出的文件名称
            string fileName = "关区管理" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
            //把Excel转为流，输出
            //创建文件流
            System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
            //将工作薄写入文件流
            excelBook.Write(bookStream);
            //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
            bookStream.Seek(0, System.IO.SeekOrigin.Begin);
            //Stream对象,文件类型,文件名称
            return File(bookStream, "application/vnd.ms-excel", fileName);
        }
        /// <summary>
        /// 关区管理报表
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievementCustoms(int GatedotID, string CustomsName, string CustomsCode)
        {
            var list = (from tbCustoms in myModels.SYS_Customs
                        join tbGatedot in myModels.SYS_Gatedot on tbCustoms.GatedotID equals tbGatedot.GatedotID
                        select new Customs
                        {
                            CustomsID = tbCustoms.CustomsID,
                            GatedotID = tbCustoms.GatedotID,
                            GatedotName = tbGatedot.GatedotName,
                            CustomsName = tbCustoms.CustomsName,
                            CustomsCode = tbCustoms.CustomsCode,
                            EnglishName = tbCustoms.EnglishName,
                            MobilePhone = tbCustoms.MobilePhone,
                            Mobile = tbCustoms.Mobile,
                            Linkman = tbCustoms.Linkman,
                            Email = tbCustoms.Email,
                            Faxes = tbCustoms.Faxes,
                            PostCode = tbCustoms.PostCode,
                            //OfficeHours = tbCustoms.OfficeHours,
                            timer = tbCustoms.OfficeHours.ToString(),
                            timerer = tbCustoms.ClosingTime.ToString(),
                            //ClosingTime = tbCustoms.ClosingTime,
                            Site = tbCustoms.Site,
                            Website = tbCustoms.Website,
                            Describe = tbCustoms.Describe,
                            WhetherStarta="",
                            WhetherStartb = tbCustoms.WhetherStart
                        }).ToList();
            if (GatedotID > 0)
            {
                list = list.Where(m => m.GatedotID == GatedotID).ToList();
            }
            if (!string.IsNullOrEmpty(CustomsName))
            {
                list = list.Where(m => m.CustomsName.Contains(CustomsName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(CustomsCode))
            {
                list = list.Where(m => m.CustomsCode.Contains(CustomsCode.Trim())).ToList();
            }
            for(int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherStartb == true)
                {
                    list[i].WhetherStarta = "√";
                }
                else if (list[i].WhetherStartb == false)
                {
                    list[i].WhetherStarta = "×";
                }
            }
            //把linq类型的数据listResult转化为DataTable类型数据
            DataTable dt = LINQToDataTableCustoms(list);
            //第一步：实例化数据集`
            PrintReport.ReportDB dbReport = new PrintReport.ReportDB();
            //第二步：将dt的数据放入数据集的数据表中
            dbReport.Tables["Customs"].Merge(dt);
            //第三步：实例化报表模板
            PrintReport.customs rp = new PrintReport.customs();
            //第四步：获取报表物理文件地址
            string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                + "Areas\\Basics\\PrintReport\\customs.rpt";
            //第五步：把报表文件加载到ReportDocument
            rp.Load(strRptPath);
            //第六步：设置报表数据源
            rp.SetDataSource(dbReport);
            //第七步：把ReportDocument转化为文件流
            Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
            return File(stream, "application/pdf");
        }
        // //将IEnumerable<T>类型的集合转换为DataTable类型
        public DataTable LINQToDataTableCustoms<T>(IEnumerable<T> varlist)
        {
            //定义要返回的DataTable对象
            DataTable dtReturn = new DataTable();
            //保存列集合的属性信息数组
            PropertyInfo[] oProps = null;
            if (varlist == null) return dtReturn;//安全性检查
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
                        //得到属性的类型
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
        #region 运输路线
        /// <summary>
        /// 查询运输路线
        /// </summary>
        /// <returns></returns>
        public ActionResult selectHaulWay(BsgridPage bsgridPage,int? MentionAreaID, int? AlsoTankAreaID, int? LoadingAreaID)
        {
            var list = (from tbHaulWay in myModels.SYS_HaulWay
                        join tbGatedot in myModels.SYS_Gatedot on tbHaulWay.LoadingAreaID equals tbGatedot.GatedotID
                        join tbCustoms in myModels.SYS_Customs on tbHaulWay.CustomsAreaID equals tbCustoms.CustomsID
                        join tbMention in myModels.SYS_Mention on tbHaulWay.MentionAreaID equals tbMention.MentionID
                        into tbMention from Mention in tbMention.DefaultIfEmpty()
                        join tbMentions in myModels.SYS_Mention on tbHaulWay.AlsoTankAreaID equals tbMentions.MentionID
                        into tbMentions from Mentions in tbMentions.DefaultIfEmpty()
                        select new HaulWay
                        {
                            HaulWayID=tbHaulWay.HaulWayID,//主表ID
                            HaulWayDescription= Mention.Abbreviation+"→"+tbGatedot.GatedotName+"→"+Mentions.Abbreviation,//运输路线
                            MentionAreaID= tbHaulWay.MentionAreaID,//提柜区域ID
                            MentionArea= Mention.Abbreviation,//提柜区域名称
                            AlsoTankAreaID= tbHaulWay.AlsoTankAreaID,//还柜区域ID
                            AlsoTankArea= Mentions.Abbreviation,//还柜区域名称
                            LoadingAreaID=tbHaulWay.LoadingAreaID,//装卸区域ID
                            GatedotName=tbGatedot.GatedotName,//装卸区域名称
                            CustomsAreaID=tbHaulWay.CustomsAreaID,//报关区域ID
                            CustomsName =tbCustoms.CustomsName,//报关区域名称
                            OutwardVoyage=tbHaulWay.OutwardVoyage,//去程(单位KM)
                            ReturnTrip =tbHaulWay.ReturnTrip,//回城(单位KM)
                            OutwardVoyageFree=tbHaulWay.OutwardVoyageFree,//去程路桥费
                            ReturnTripFree=tbHaulWay.ReturnTripFree,//回城路桥费
                            OutwardVoyageTime=tbHaulWay.OutwardVoyageTime,//去程预计时间(H)
                            OutwardVoyageTimetimer= tbHaulWay.OutwardVoyageTime.ToString(),
                            ReturnTripTime =tbHaulWay.ReturnTripTime,//回程预计时间(H)
                            ReturnTripTimetimer= tbHaulWay.ReturnTripTime.ToString(),
                            TotalTime =tbHaulWay.TotalTime,//总预计时间(H)
                            TotalTimetimer= tbHaulWay.TotalTime.ToString(),
                            WhetherStart =tbHaulWay.WhetherStart,//是否启用
                            Describe=tbHaulWay.Describe//描述
                        }).ToList();
            if (MentionAreaID > 0)
            {
                list = list.Where(m => m.MentionAreaID == MentionAreaID).ToList();
            }
            if (AlsoTankAreaID > 0)
            {
                list = list.Where(m => m.AlsoTankAreaID == AlsoTankAreaID).ToList();
            }
            if (LoadingAreaID > 0)
            {
                list = list.Where(m => m.LoadingAreaID == LoadingAreaID).ToList();
            }
            int totalRow = list.Count();
            List<HaulWay> notices = list.OrderByDescending(p => p.HaulWayID).
              Skip(bsgridPage.GetStartIndex()).
              Take(bsgridPage.pageSize).
              ToList();
            Bsgrid<HaulWay> bsgrid = new Bsgrid<HaulWay>();
            bsgrid.success = true;
            bsgrid.totalRows = totalRow;
            bsgrid.curPage = bsgridPage.curPage;
            bsgrid.data = notices;
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 新增运输路线
        /// </summary>
        /// <returns></returns>
        public ActionResult InsertHaulWay(SYS_HaulWay tbhaulWay)
        {
            string strMsg = "fail";
            try
            {
                //int oldCustoms = (from tbHaulWay in myModels.SYS_HaulWay
                //                  where tbHaulWay.HaulWayDescription == tbhaulWay.HaulWayDescription
                //                  select tbHaulWay).Count();
                //if (oldCustoms == 0)
                //{
                    myModels.SYS_HaulWay.Add(tbhaulWay);
                    if (myModels.SaveChanges() > 0)
                    {
                        strMsg = "数据新增成功！";
                    }
                //}
                //else
                //{
                //    strMsg = "已有运输路线描述，不能新增！";
                //}
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 修改运输路线
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateHaulWay(SYS_HaulWay pwhaulWay)
        {
            string strMsg = "fail";
            try
            {
                SYS_HaulWay dbHaulWay = (from tbHaulWay in myModels.SYS_HaulWay
                                         where tbHaulWay.HaulWayID == pwhaulWay.HaulWayID
                                         select tbHaulWay).Single();
                dbHaulWay.HaulWayDescription = pwhaulWay.HaulWayDescription;
                dbHaulWay.MentionAreaID = pwhaulWay.MentionAreaID;
                dbHaulWay.LoadingAreaID = pwhaulWay.LoadingAreaID;
                dbHaulWay.CustomsAreaID = pwhaulWay.CustomsAreaID;
                dbHaulWay.AlsoTankAreaID = pwhaulWay.AlsoTankAreaID;
                dbHaulWay.OutwardVoyage = pwhaulWay.OutwardVoyage;
                dbHaulWay.ReturnTrip = pwhaulWay.ReturnTrip;
                dbHaulWay.OutwardVoyageFree = pwhaulWay.OutwardVoyageFree;
                dbHaulWay.ReturnTripFree = pwhaulWay.ReturnTripFree;
                dbHaulWay.OutwardVoyageTime = pwhaulWay.OutwardVoyageTime;
                dbHaulWay.ReturnTripTime = pwhaulWay.ReturnTripTime;
                dbHaulWay.TotalTime = pwhaulWay.TotalTime;
                dbHaulWay.WhetherStart = pwhaulWay.WhetherStart;
                dbHaulWay.Describe = pwhaulWay.Describe;
                myModels.Entry(dbHaulWay).State = System.Data.Entity.EntityState.Modified;
                if (myModels.SaveChanges() > 0)
                {
                    strMsg = "修改成功！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 回填运输路线
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectHaulWayById(int HaulWayID)
        {
            try
            {
                var list = (from tbHaulWay in myModels.SYS_HaulWay
                            join tbGatedot in myModels.SYS_Gatedot on tbHaulWay.LoadingAreaID equals tbGatedot.GatedotID
                            join tbCustoms in myModels.SYS_Customs on tbHaulWay.CustomsAreaID equals tbCustoms.CustomsID
                            join tbMention in myModels.SYS_Mention on tbHaulWay.MentionAreaID equals tbMention.MentionID
                            into tbMention
                            from Mention in tbMention.DefaultIfEmpty()
                            join tbMentions in myModels.SYS_Mention on tbHaulWay.AlsoTankAreaID equals tbMentions.MentionID
                            into tbMentions
                            from Mentions in tbMentions.DefaultIfEmpty()
                            where tbHaulWay.HaulWayID== HaulWayID
                            select new HaulWay
                            {
                                HaulWayID = tbHaulWay.HaulWayID,//主表ID
                                HaulWayDescription = Mention.Abbreviation + "-" + tbGatedot.GatedotName + "-" + Mentions.Abbreviation,//运输路线描述
                                MentionAreaID = tbHaulWay.MentionAreaID,//提柜区域ID
                                MentionArea = Mention.Abbreviation.Trim(),//提柜区域名称
                                AlsoTankAreaID = tbHaulWay.AlsoTankAreaID,//还柜区域ID
                                AlsoTankArea = Mentions.Abbreviation.Trim(),//还柜区域名称
                                LoadingAreaID = tbHaulWay.LoadingAreaID,//装卸区域ID
                                GatedotName = tbGatedot.GatedotName.Trim(),//装卸区域名称
                                CustomsAreaID = tbHaulWay.CustomsAreaID,//报关区域ID
                                CustomsName = tbCustoms.CustomsName.Trim(),//报关区域名称
                                OutwardVoyage = tbHaulWay.OutwardVoyage,//去程(单位KM)
                                ReturnTrip = tbHaulWay.ReturnTrip,//回城(单位KM)
                                OutwardVoyageFree = tbHaulWay.OutwardVoyageFree,//去程路桥费
                                ReturnTripFree = tbHaulWay.ReturnTripFree,//回城路桥费
                                OutwardVoyageTime = tbHaulWay.OutwardVoyageTime,//去程预计时间(H)
                                ReturnTripTime = tbHaulWay.ReturnTripTime,//回程预计时间(H)
                                TotalTime = tbHaulWay.TotalTime,//总预计时间(H)
                                WhetherStart = tbHaulWay.WhetherStart,//是否启用
                                Describe = tbHaulWay.Describe//描述
                            }).Single();
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return null;
            }
        }
        /// <summary>
        /// 删除运输路线
        /// </summary>
        /// <returns></returns>
        public ActionResult DelectHaulWay(int HaulWayID)
        {
            string strMsg = "fail";
            try
            {
                var listHaulWay = myModels.SYS_HaulWay.Where(m => m.HaulWayID == HaulWayID).Single();
                myModels.SYS_HaulWay.Remove(listHaulWay);
                if (myModels.SaveChanges() > 0)
                {
                    strMsg = "true";
                }
            }
            catch (Exception e)
            {
                strMsg = "fail";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 导出运输路线
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportToExcelHaulWay(int MentionAreaID, int AlsoTankAreaID, int LoadingAreaID)
        {
            var list = (from tbHaulWay in myModels.SYS_HaulWay
                        join tbGatedot in myModels.SYS_Gatedot on tbHaulWay.LoadingAreaID equals tbGatedot.GatedotID
                        join tbCustoms in myModels.SYS_Customs on tbHaulWay.CustomsAreaID equals tbCustoms.CustomsID
                        join tbMention in myModels.SYS_Mention on tbHaulWay.MentionAreaID equals tbMention.MentionID
                        into tbMention
                        from Mention in tbMention.DefaultIfEmpty()
                        join tbMentions in myModels.SYS_Mention on tbHaulWay.AlsoTankAreaID equals tbMentions.MentionID
                        into tbMentions
                        from Mentions in tbMentions.DefaultIfEmpty()
                        select new HaulWay
                        {
                            HaulWayID = tbHaulWay.HaulWayID,//主表ID
                            HaulWayDescription = Mention.Abbreviation + "-" + tbGatedot.GatedotName + "-" + Mentions.Abbreviation,//运输路线描述
                            MentionAreaID = tbHaulWay.MentionAreaID,//提柜区域ID
                            MentionArea = Mention.Abbreviation,//提柜区域名称
                            AlsoTankAreaID = tbHaulWay.AlsoTankAreaID,//还柜区域ID
                            AlsoTankArea = Mentions.Abbreviation,//还柜区域名称
                            LoadingAreaID = tbHaulWay.LoadingAreaID,//装卸区域ID
                            GatedotName = tbGatedot.GatedotName,//装卸区域名称
                            CustomsAreaID = tbHaulWay.CustomsAreaID,//报关区域ID
                            CustomsName = tbCustoms.CustomsName,//报关区域名称
                            OutwardVoyage = tbHaulWay.OutwardVoyage,//去程(单位KM)
                            ReturnTrip = tbHaulWay.ReturnTrip,//回城(单位KM)
                            OutwardVoyageFree = tbHaulWay.OutwardVoyageFree,//去程路桥费
                            ReturnTripFree = tbHaulWay.ReturnTripFree,//回城路桥费
                            OutwardVoyageTime = tbHaulWay.OutwardVoyageTime,//去程预计时间(H)
                            ReturnTripTime = tbHaulWay.ReturnTripTime,//回程预计时间(H)
                            TotalTime = tbHaulWay.TotalTime,//总预计时间(H)
                            WhetherStarta = "",
                            WhetherStartb = tbHaulWay.WhetherStart,//是否启用
                            Describe = tbHaulWay.Describe//描述
                        }).ToList();
            if (MentionAreaID > 0)
            {
                list = list.Where(m => m.MentionAreaID == MentionAreaID).ToList();
            }
            if (AlsoTankAreaID > 0)
            {
                list = list.Where(m => m.AlsoTankAreaID == AlsoTankAreaID).ToList();
            }
            if (LoadingAreaID > 0)
            {
                list = list.Where(m => m.LoadingAreaID == LoadingAreaID).ToList();
            }
            for(int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherStartb == true)
                {
                    list[i].WhetherStarta = "启用";
                }
                else if (list[i].WhetherStartb == false)
                {
                    list[i].WhetherStarta = "不启用";
                }
            }
            //查询数据
            List<HaulWay> listExaminee = list.ToList();
            //二：代码创建一个Excel表格（这里称为工作簿）
            //创建Excel文件的对象 工作簿(调用NPOI文件)
            HSSFWorkbook excelBook = new HSSFWorkbook();
            //创建Excel工作表 Sheet
            ISheet sheet1 = excelBook.CreateSheet("运输路线");
            //给Sheet添加第一行的头部标题
            IRow row1 = sheet1.CreateRow(0);
            //给标题的每一个单元格赋值
            row1.CreateCell(0).SetCellValue("运输路线描述");
            row1.CreateCell(1).SetCellValue("提柜区域");
            row1.CreateCell(2).SetCellValue("装卸区域");
            row1.CreateCell(3).SetCellValue("报关区域");
            row1.CreateCell(4).SetCellValue("还柜区域");
            row1.CreateCell(5).SetCellValue("去程(单位KM)");
            row1.CreateCell(6).SetCellValue("回城(单位KM)");
            row1.CreateCell(7).SetCellValue("去程路桥费");
            row1.CreateCell(8).SetCellValue("回城路桥费");
            row1.CreateCell(9).SetCellValue("去程预计时间(H)");
            row1.CreateCell(10).SetCellValue("回程预计时间(H)");
            row1.CreateCell(11).SetCellValue("总预计时间(H)");
            row1.CreateCell(12).SetCellValue("描述");
            row1.CreateCell(13).SetCellValue("是否启用");
            //添加数据行：将表格数据逐步写入sheet1各个行中（也就是给每一个单元格赋值）
            for (int i = 0; i < listExaminee.Count; i++)
            {
                //创建行
                IRow rowTemp = sheet1.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listExaminee[i].HaulWayDescription);
                rowTemp.CreateCell(1).SetCellValue(listExaminee[i].MentionArea);
                rowTemp.CreateCell(2).SetCellValue(listExaminee[i].GatedotName);
                rowTemp.CreateCell(3).SetCellValue(listExaminee[i].CustomsName);
                rowTemp.CreateCell(4).SetCellValue(listExaminee[i].AlsoTankArea);
                rowTemp.CreateCell(5).SetCellValue(listExaminee[i].OutwardVoyage.ToString());
                rowTemp.CreateCell(6).SetCellValue(listExaminee[i].ReturnTrip.ToString());
                rowTemp.CreateCell(7).SetCellValue(listExaminee[i].OutwardVoyageFree.ToString());
                rowTemp.CreateCell(8).SetCellValue(listExaminee[i].ReturnTripFree.ToString());
                rowTemp.CreateCell(9).SetCellValue(listExaminee[i].OutwardVoyageTime.ToString());
                rowTemp.CreateCell(10).SetCellValue(listExaminee[i].ReturnTripTime.ToString());
                rowTemp.CreateCell(11).SetCellValue(listExaminee[i].TotalTime.ToString());
                rowTemp.CreateCell(12).SetCellValue(listExaminee[i].Describe);
                rowTemp.CreateCell(13).SetCellValue(listExaminee[i].WhetherStarta);
            }
            //输出的文件名称
            string fileName = "运输路线" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
            //把Excel转为流，输出
            //创建文件流
            System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
            //将工作薄写入文件流
            excelBook.Write(bookStream);
            //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
            bookStream.Seek(0, System.IO.SeekOrigin.Begin);
            //Stream对象,文件类型,文件名称
            return File(bookStream, "application/vnd.ms-excel", fileName);
        }
        /// <summary>
        /// 运输路线报表
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievementHaulWay(int MentionAreaID, int AlsoTankAreaID, int LoadingAreaID)
        {
            var list = (from tbHaulWay in myModels.SYS_HaulWay
                        join tbGatedot in myModels.SYS_Gatedot on tbHaulWay.LoadingAreaID equals tbGatedot.GatedotID
                        join tbCustoms in myModels.SYS_Customs on tbHaulWay.CustomsAreaID equals tbCustoms.CustomsID
                        join tbMention in myModels.SYS_Mention on tbHaulWay.MentionAreaID equals tbMention.MentionID
                        into tbMention
                        from Mention in tbMention.DefaultIfEmpty()
                        join tbMentions in myModels.SYS_Mention on tbHaulWay.AlsoTankAreaID equals tbMentions.MentionID
                        into tbMentions
                        from Mentions in tbMentions.DefaultIfEmpty()
                        select new HaulWay
                        {
                            HaulWayID = tbHaulWay.HaulWayID,//主表ID
                            HaulWayDescription = Mention.Abbreviation + "-" + tbGatedot.GatedotName + "-" + Mentions.Abbreviation,//运输路线描述
                            MentionAreaID = tbHaulWay.MentionAreaID,//提柜区域ID
                            MentionArea = Mention.Abbreviation,//提柜区域名称
                            AlsoTankAreaID = tbHaulWay.AlsoTankAreaID,//还柜区域ID
                            AlsoTankArea = Mentions.Abbreviation,//还柜区域名称
                            LoadingAreaID = tbHaulWay.LoadingAreaID,//装卸区域ID
                            GatedotName = tbGatedot.GatedotName,//装卸区域名称
                            CustomsAreaID = tbHaulWay.CustomsAreaID,//报关区域ID
                            CustomsName = tbCustoms.CustomsName,//报关区域名称
                            OutwardVoyage = tbHaulWay.OutwardVoyage,//去程(单位KM)
                            ReturnTrip = tbHaulWay.ReturnTrip,//回城(单位KM)
                            OutwardVoyageFree = tbHaulWay.OutwardVoyageFree,//去程路桥费
                            ReturnTripFree = tbHaulWay.ReturnTripFree,//回城路桥费
                            OutwardVoyageTime = tbHaulWay.OutwardVoyageTime,//去程预计时间(H)
                            ReturnTripTime = tbHaulWay.ReturnTripTime,//回程预计时间(H)
                            TotalTime = tbHaulWay.TotalTime,//总预计时间(H)
                            WhetherStarta = "",
                            WhetherStartb = tbHaulWay.WhetherStart,//是否启用
                            Describe = tbHaulWay.Describe//描述
                        }).ToList();
            if (MentionAreaID > 0)
            {
                list = list.Where(m => m.MentionAreaID == MentionAreaID).ToList();
            }
            if (AlsoTankAreaID > 0)
            {
                list = list.Where(m => m.AlsoTankAreaID == AlsoTankAreaID).ToList();
            }
            if (LoadingAreaID > 0)
            {
                list = list.Where(m => m.LoadingAreaID == LoadingAreaID).ToList();
            }
            for(int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherStartb == true)
                {
                    list[i].WhetherStarta = "√";
                }
                else if (list[i].WhetherStartb == false)
                {
                    list[i].WhetherStarta = "×";
                }
            }
            //把linq类型的数据listResult转化为DataTable类型数据
            DataTable dt = LINQToDataTableHaulWay(list);
            //第一步：实例化数据集`
            PrintReport.ReportDB dbReport = new PrintReport.ReportDB();
            //第二步：将dt的数据放入数据集的数据表中
            dbReport.Tables["HaulWay"].Merge(dt);
            //第三步：实例化报表模板
            PrintReport.haulWay rp = new PrintReport.haulWay();
            //第四步：获取报表物理文件地址
            string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                + "Areas\\Basics\\PrintReport\\haulWay.rpt";
            //第五步：把报表文件加载到ReportDocument
            rp.Load(strRptPath);
            //第六步：设置报表数据源
            rp.SetDataSource(dbReport);
            //第七步：把ReportDocument转化为文件流
            Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
            return File(stream, "application/pdf");
        }
        // //将IEnumerable<T>类型的集合转换为DataTable类型
        public DataTable LINQToDataTableHaulWay<T>(IEnumerable<T> varlist)
        {
            //定义要返回的DataTable对象
            DataTable dtReturn = new DataTable();
            //保存列集合的属性信息数组
            PropertyInfo[] oProps = null;
            if (varlist == null) return dtReturn;//安全性检查
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
                        //得到属性的类型
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
        #region 司机产值
        /// <summary>
        /// 查询司机报表
        /// </summary>
        /// <returns></returns>
        public ActionResult selectChauffeurOffer(BsgridPage bsgridPage,string Begin,string Finish,string Always)
        {
            var list = (from tbOffer in myModels.SYS_Offer
                        join tbClient in myModels.SYS_Client on tbOffer.ClientID equals tbClient.ClientID
                        where tbOffer.OfferType.Trim() == "司机产值"
                        select new chauffeuroffer
                        {
                            ClientID = tbOffer.ClientID,
                            ChineseName = tbClient.ChineseName,
                            OfferID = tbOffer.OfferID,
                            OfferType = tbOffer.OfferType,
                            OfferDate = tbOffer.OfferDate,
                            OfferDatetime = tbOffer.OfferDate.ToString(),
                            TakeEffectDate = tbOffer.TakeEffectDate,
                            TakeEffectDatetime = tbOffer.TakeEffectDate.ToString(),
                            LoseEfficacyDate = tbOffer.LoseEfficacyDate,
                            LoseEfficacyDatetime = tbOffer.LoseEfficacyDate.ToString(),
                            WhetherShui = tbOffer.WhetherShui,
                            Remark = tbOffer.Remark
                        }).ToList();
            if (!string.IsNullOrEmpty(Begin) && !string.IsNullOrEmpty(Finish))
            {
                DateTime BeginDatetime = Convert.ToDateTime(Begin);
                DateTime FinishDatetime = Convert.ToDateTime(Finish);
                list = list.Where(m => m.TakeEffectDate >= BeginDatetime && m.LoseEfficacyDate <= FinishDatetime).ToList();
            }
            if (!string.IsNullOrEmpty(Always))
            {
                DateTime AlwaysDatetime = Convert.ToDateTime(Always);
                list = list.Where(m => m.OfferDate <= AlwaysDatetime || m.OfferDate >= AlwaysDatetime).ToList();
            }
            int totalRow = list.Count();
            List<chauffeuroffer> notices = list.OrderByDescending(p => p.OfferID).
              Skip(bsgridPage.GetStartIndex()).
              Take(bsgridPage.pageSize).
              ToList();
            Bsgrid<chauffeuroffer> bsgrid = new Bsgrid<chauffeuroffer>();
            bsgrid.success = true;
            bsgrid.totalRows = totalRow;
            bsgrid.curPage = bsgridPage.curPage;
            bsgrid.data = notices;
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 查询司机明细
        /// </summary>
        /// <returns></returns>
        public ActionResult selectChauffeurOfferDetail(BsgridPage bsgridPage,int OfferID,int StaffID,string Currency,string EtryClasses,string CabinetType)
        {
            var list = (from tbOfferDetail in myModels.SYS_OfferDetail
                        join tbOfferID in myModels.SYS_Offer on tbOfferDetail.OfferID equals tbOfferID.OfferID
                        join tbExpenseID in myModels.SYS_Expense on tbOfferDetail.ExpenseID equals tbExpenseID.ExpenseID
                        join tbHaulWayID in myModels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWayID.HaulWayID
                        join tbStaffID in myModels.SYS_Staff on tbOfferDetail.StaffID equals tbStaffID.StaffID
                        where tbOfferDetail.OfferID== OfferID
                        select new chauffeurOfferDetail
                        {
                            OfferID=tbOfferDetail.OfferID,
                            OfferType=tbOfferID.OfferType,
                            ExpenseID =tbOfferDetail.ExpenseID,
                            ExpenseName=tbExpenseID.ExpenseName,
                            HaulWayID =tbOfferDetail.HaulWayID,
                            HaulWayDescription=tbHaulWayID.HaulWayDescription,
                            StaffID=tbOfferDetail.StaffID,
                            StaffName=tbStaffID.StaffName,
                            OfferDetailID =tbOfferDetail.OfferDetailID,
                            EtryClasses= tbOfferDetail.EtryClasses,
                            CabinetType=tbOfferDetail.CabinetType,
                            Money=tbOfferDetail.Money,
                            Remark=tbOfferDetail.Remark,
                            WorkCategory=tbOfferDetail.WorkCategory,
                            BoxQuantity=tbOfferDetail.BoxQuantity,
                            Currency=tbOfferDetail.Currency
                        }).ToList();
            if (StaffID > 0)
            {
                list = list.Where(m => m.StaffID == StaffID).ToList();
            }
            if (!string.IsNullOrEmpty(Currency))
            {
                list = list.Where(m => m.Currency.Contains(Currency.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(EtryClasses))
            {
                list = list.Where(m => m.EtryClasses.Contains(EtryClasses.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(CabinetType))
            {
                list = list.Where(m => m.CabinetType.Contains(CabinetType.Trim())).ToList();
            }
            int totalRow = list.Count();
            List<chauffeurOfferDetail> notices = list.OrderByDescending(p => p.OfferDetailID).
              Skip(bsgridPage.GetStartIndex()).
              Take(bsgridPage.pageSize).
              ToList();
            Bsgrid<chauffeurOfferDetail> bsgrid = new Bsgrid<chauffeurOfferDetail>();
            bsgrid.success = true;
            bsgrid.totalRows = totalRow;
            bsgrid.curPage = bsgridPage.curPage;
            bsgrid.data = notices;
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        #endregion
        #region 费用项目
        /// <summary>
        /// 查询费用项目
        /// </summary>
        /// <returns></returns>
        public ActionResult selectExpense(BsgridPage bsgridPage, string ExpenseCode,string ExpenseName,string SettleAccounts)
        {
            var list = (from tbExpense in myModels.SYS_Expense
                        orderby tbExpense.CostSortingCode ascending
                        select new Expense
                        {
                            ExpenseID=tbExpense.ExpenseID,
                            ExpenseCode=tbExpense.ExpenseCode,
                            ExpenseName=tbExpense.ExpenseName,
                            SettleAccounts=tbExpense.SettleAccounts,
                            CostAttributes = tbExpense.CostAttributes,
                            CostSortingCode =tbExpense.CostSortingCode,
                            WhetherStart=tbExpense.WhetherStart,
                            Remark=tbExpense.Remark
                        }).ToList();
            if (!string.IsNullOrEmpty(ExpenseCode))
            {
                list = list.Where(s => s.ExpenseCode.Contains(ExpenseCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ExpenseName))
            {
                list = list.Where(s => s.ExpenseName.Contains(ExpenseName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(SettleAccounts))
            {
                list = list.Where(s => s.SettleAccounts.Contains(SettleAccounts.Trim())).ToList();
            }
            int totalRow = list.Count();
            List<Expense> notices = list.
              Skip(bsgridPage.GetStartIndex()).
              Take(bsgridPage.pageSize).
              ToList();
            Bsgrid<Expense> bsgrid = new Bsgrid<Expense>();
            bsgrid.success = true;
            bsgrid.totalRows = totalRow;
            bsgrid.curPage = bsgridPage.curPage;
            bsgrid.data = notices;
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 新增费用项目
        /// </summary>
        /// <returns></returns>
        public ActionResult InsertExpense(SYS_Expense tbexpense)
        {
            string strMsg = "fail";
            try
            {
                int oldExpense = (from tbExpense in myModels.SYS_Expense
                               where tbExpense.ExpenseCode == tbexpense.ExpenseCode
                               select tbExpense).Count();
                if (oldExpense == 0)
                {
                    int oldname = (from tbExpense in myModels.SYS_Expense
                                   where tbExpense.ExpenseName == tbexpense.ExpenseName
                                   select tbExpense).Count();
                    if (oldname == 0)
                    {
                        myModels.SYS_Expense.Add(tbexpense);
                        if (myModels.SaveChanges() > 0)
                        {
                            strMsg = "数据新增成功！";
                        }
                    }
                    else
                    {
                        strMsg = "已有费用名称！";
                    }
                }
                else
                {
                    strMsg = "已有费用编码！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 修改费用项目
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateExpense(SYS_Expense pwexpense)
        {
            string strMsg = "fail";
            try
            {
                if (myModels.SYS_Expense.Where(m => m.ExpenseCode == pwexpense.ExpenseCode.Trim() && m.ExpenseID != pwexpense.ExpenseID).Count() == 0)
                {
                    if (myModels.SYS_Expense.Where(m => m.ExpenseName == pwexpense.ExpenseName.Trim() && m.ExpenseID != pwexpense.ExpenseID).Count() == 0)
                    {
                        var list = (from tb in myModels.SYS_Expense
                                    where tb.CostSortingCode != 0
                                    select tb).Count();
                        if (list != 0)
                        {
                            SYS_Expense dbexpense = (from tbExpense in myModels.SYS_Expense
                                                     where tbExpense.ExpenseID == pwexpense.ExpenseID
                                                     select tbExpense).Single();
                            dbexpense.ExpenseCode = pwexpense.ExpenseCode;
                            dbexpense.ExpenseName = pwexpense.ExpenseName;
                            dbexpense.SettleAccounts = pwexpense.SettleAccounts;
                            dbexpense.CostAttributes = pwexpense.CostAttributes;
                            dbexpense.CostSortingCode = pwexpense.CostSortingCode;
                            dbexpense.WhetherStart = pwexpense.WhetherStart;
                            dbexpense.Remark = pwexpense.Remark;
                            myModels.Entry(dbexpense).State = System.Data.Entity.EntityState.Modified;
                            if (myModels.SaveChanges() > 0)
                            {
                                strMsg = "修改成功！";
                            }
                        }
                        else
                        {
                            strMsg = "费用排序码为0，不能进行修改！";
                        }
                    }
                    else
                    {
                        strMsg = "已有费用名称，无法修改！";
                    }
                }
                else
                {
                    strMsg = "已有费用编码，无法修改！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 回填费用项目
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectExpenseById(int ExpenseID)
        {
            try
            {
                var list = (from tbExpense in myModels.SYS_Expense
                            where tbExpense.ExpenseID == ExpenseID
                            select new Expense
                            {
                                ExpenseID = tbExpense.ExpenseID,
                                ExpenseCode = tbExpense.ExpenseCode,
                                ExpenseName = tbExpense.ExpenseName,
                                SettleAccounts = tbExpense.SettleAccounts.Trim(),
                                CostAttributes = tbExpense.CostAttributes.Trim(),
                                CostSortingCode = tbExpense.CostSortingCode,
                                WhetherStart = tbExpense.WhetherStart,
                                Remark = tbExpense.Remark
                            }).Single();
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return null;
            }

        }
        /// <summary>
        /// 删除费用项目
        /// </summary>
        /// <returns></returns>
        public ActionResult DelectExpense(int ExpenseID)
        {
            string strMsg = "fail";
            try
            {
                var listExpense = myModels.SYS_Expense.Where(m => m.ExpenseID == ExpenseID).Single();
                myModels.SYS_Expense.Remove(listExpense);
                if (myModels.SaveChanges() > 0)
                {
                    strMsg = "删除成功！";
                }
            }
            catch (Exception e)
            {
                strMsg = "fail";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 导出费用项目
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportToExcelExpense(string ExpenseCode, string ExpenseName, string SettleAccounts)
        {
            var list = (from tbExpense in myModels.SYS_Expense
                        select new Expense
                        {
                            ExpenseID = tbExpense.ExpenseID,
                            ExpenseCode = tbExpense.ExpenseCode,
                            ExpenseName = tbExpense.ExpenseName,
                            SettleAccounts = tbExpense.SettleAccounts,
                            CostAttributes = tbExpense.CostAttributes,
                            CostSortingCode = tbExpense.CostSortingCode,
                            WhetherStarta = "",
                            WhetherStartl = tbExpense.WhetherStart,
                            Remark = tbExpense.Remark
                        }).ToList();
            if (!string.IsNullOrEmpty(ExpenseCode))
            {
                list = list.Where(s => s.ExpenseCode.Contains(ExpenseCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ExpenseName))
            {
                list = list.Where(s => s.ExpenseName.Contains(ExpenseName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(SettleAccounts))
            {
                list = list.Where(s => s.SettleAccounts.Contains(SettleAccounts.Trim())).ToList();
            }
            for(int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherStartl == true)
                {
                    list[i].WhetherStarta = "启用";
                }
                else if(list[i].WhetherStartl==false)
                {
                    list[i].WhetherStarta = "不启用";
                }
            }
            //查询数据
            List<Expense> listExaminee = list.ToList();
            //二：代码创建一个Excel表格（这里称为工作簿）
            //创建Excel文件的对象 工作簿(调用NPOI文件)
            HSSFWorkbook excelBook = new HSSFWorkbook();
            //创建Excel工作表 Sheet
            ISheet sheet1 = excelBook.CreateSheet("费用项目");
            //给Sheet添加第一行的头部标题
            IRow row1 = sheet1.CreateRow(0);
            //给标题的每一个单元格赋值
            row1.CreateCell(0).SetCellValue("费用编码");
            row1.CreateCell(1).SetCellValue("费用项目名称");
            row1.CreateCell(2).SetCellValue("费用子类型");
            row1.CreateCell(3).SetCellValue("费用属性");
            row1.CreateCell(4).SetCellValue("费用排序码");
            row1.CreateCell(5).SetCellValue("是否启用");
            row1.CreateCell(6).SetCellValue("备注");
            //添加数据行：将表格数据逐步写入sheet1各个行中（也就是给每一个单元格赋值）
            for (int i = 0; i < listExaminee.Count; i++)
            {
                //创建行
                IRow rowTemp = sheet1.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listExaminee[i].ExpenseCode);
                rowTemp.CreateCell(1).SetCellValue(listExaminee[i].ExpenseName);
                rowTemp.CreateCell(2).SetCellValue(listExaminee[i].SettleAccounts);
                rowTemp.CreateCell(3).SetCellValue(listExaminee[i].CostAttributes);
                rowTemp.CreateCell(4).SetCellValue(listExaminee[i].CostSortingCode.ToString());
                rowTemp.CreateCell(5).SetCellValue(listExaminee[i].WhetherStarta);
                rowTemp.CreateCell(6).SetCellValue(listExaminee[i].Remark);
            }
            //输出的文件名称
            string fileName = "费用项目" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
            //把Excel转为流，输出
            //创建文件流
            System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
            //将工作薄写入文件流
            excelBook.Write(bookStream);
            //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
            bookStream.Seek(0, System.IO.SeekOrigin.Begin);
            //Stream对象,文件类型,文件名称
            return File(bookStream, "application/vnd.ms-excel", fileName);
        }
        /// <summary>
        /// 费用项目报表
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievementExpense(string ExpenseCode, string ExpenseName, string SettleAccounts)
        {
            var list = (from tbExpense in myModels.SYS_Expense
                        select new Expense
                        {
                            ExpenseID = tbExpense.ExpenseID,
                            ExpenseCode = tbExpense.ExpenseCode,
                            ExpenseName = tbExpense.ExpenseName,
                            SettleAccounts = tbExpense.SettleAccounts,
                            CostAttributes = tbExpense.CostAttributes,
                            CostSortingCode = tbExpense.CostSortingCode,
                            WhetherStarta = "",
                            WhetherStartl = tbExpense.WhetherStart,
                            Remark = tbExpense.Remark
                        }).ToList();
            if (!string.IsNullOrEmpty(ExpenseCode))
            {
                list = list.Where(s => s.ExpenseCode.Contains(ExpenseCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ExpenseName))
            {
                list = list.Where(s => s.ExpenseName.Contains(ExpenseName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(SettleAccounts))
            {
                list = list.Where(s => s.SettleAccounts.Contains(SettleAccounts.Trim())).ToList();
            }
            for(int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherStartl == true)
                {
                    list[i].WhetherStarta = "√";
                }
                else if (list[i].WhetherStartl == false)
                {
                    list[i].WhetherStarta = "×";
                }
            }
            //把linq类型的数据listResult转化为DataTable类型数据
            DataTable dt = LINQToDataTableExpense(list);
            //第一步：实例化数据集`
            PrintReport.ReportDB dbReport = new PrintReport.ReportDB();
            //第二步：将dt的数据放入数据集的数据表中
            dbReport.Tables["Expense"].Merge(dt);
            //第三步：实例化报表模板
            PrintReport.expense rp = new PrintReport.expense();
            //第四步：获取报表物理文件地址     
            string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                + "Areas\\Basics\\PrintReport\\expense.rpt";
            //第五步：把报表文件加载到ReportDocument
            rp.Load(strRptPath);
            //第六步：设置报表数据源
            rp.SetDataSource(dbReport);
            //第七步：把ReportDocument转化为文件流
            Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
            return File(stream, "application/pdf");
        }
        // //将IEnumerable<T>类型的集合转换为DataTable类型
        public DataTable LINQToDataTableExpense<T>(IEnumerable<T> varlist)
        {
            //定义要返回的DataTable对象
            DataTable dtReturn = new DataTable();
            //保存列集合的属性信息数组
            PropertyInfo[] oProps = null;
            if (varlist == null) return dtReturn;//安全性检查
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
                        //得到属性的类型
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
        #region  系统汇率
        /// <summary>
        /// 查询系统汇率
        /// </summary>
        /// <returns></returns>
        public ActionResult selectParities(BsgridPage bsgridPage)
        {
            var list = (from tbParities in myModels.SYS_Parities
                        select new Parities
                        {
                            ParitiesID = tbParities.ParitiesID,
                            Currency = tbParities.Currency,
                            BasicCurrency = tbParities.BasicCurrency,
                            ConversionRate = tbParities.ConversionRate
                        }).ToList();
            int totalRow = list.Count();
            List<Parities> notices = list.OrderByDescending(p => p.ParitiesID).
              Skip(bsgridPage.GetStartIndex()).
              Take(bsgridPage.pageSize).
              ToList();
            Bsgrid<Parities> bsgrid = new Bsgrid<Parities>();
            bsgrid.success = true;
            bsgrid.totalRows = totalRow;
            bsgrid.curPage = bsgridPage.curPage;
            bsgrid.data = notices;
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 新增系统汇率
        /// </summary>
        /// <returns></returns>
        public ActionResult InsertParities(SYS_Parities tbParities)
        {
            string strMsg = "fail";
            try
            {
                //int oldparities = (from tbparities in myModels.SYS_Parities
                //                   where tbparities.Currency.Trim() == tbParities.Currency.Trim() && tbparities.BasicCurrency.Trim() == tbParities.BasicCurrency.Trim()
                //                   select tbparities).Count();
                //if (oldparities == 0)
                //{
                    if (tbParities.Currency == tbParities.BasicCurrency)
                    {
                        tbParities.ConversionRate = 1;
                        myModels.SYS_Parities.Add(tbParities);
                        if (myModels.SaveChanges() > 0)
                        {
                            strMsg = "数据新增成功！";
                        }
                    }
                //}
                //else
                //{
                //    strMsg = "新增的交易币和基本币在数据库存在！";
                //}
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 修改系统汇率
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateParities(SYS_Parities pwParities)
        {
            string strMsg = "fail";
            try
            {
                SYS_Parities dbParities = (from tbparities in myModels.SYS_Parities
                                   where tbparities.ParitiesID == pwParities.ParitiesID
                                   select tbparities).Single();
                dbParities.Currency = pwParities.Currency;
                dbParities.ConversionRate = pwParities.ConversionRate;
                dbParities.BasicCurrency = pwParities.BasicCurrency;
                myModels.Entry(dbParities).State = System.Data.Entity.EntityState.Modified;
                if (myModels.SaveChanges() > 0)
                {
                    strMsg = "修改成功！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 修改回填汇率
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectParitiesById(int ParitiesID)
        {
            try
            {
                var list = (from tbParities in myModels.SYS_Parities
                            where tbParities.ParitiesID == ParitiesID
                            select new Parities
                            {
                                ParitiesID = tbParities.ParitiesID,
                                Currency = tbParities.Currency.Trim(),
                                BasicCurrency = tbParities.BasicCurrency.Trim(),
                                ConversionRate = tbParities.ConversionRate
                            }).Single();
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return null;
            }
        }
        /// <summary>
        /// 删除系统汇率
        /// </summary>
        /// <returns></returns>
        public ActionResult DelectParities(int ParitiesID)
        {
            string strMsg = "fail";
            try
            {
                var listParities = myModels.SYS_Parities.Where(m => m.ParitiesID == ParitiesID).Single();
                myModels.SYS_Parities.Remove(listParities);
                if (myModels.SaveChanges() > 0)
                {
                    strMsg = "true";
                }
            }
            catch (Exception e)
            {
                strMsg = "fail";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 导出系统汇率
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportToExcelParities(string Currency, string BasicCurrency)
        {
            var list = (from tbParities in myModels.SYS_Parities
                        select new Parities
                        {
                            ParitiesID = tbParities.ParitiesID,
                            Currency = tbParities.Currency,
                            BasicCurrency = tbParities.BasicCurrency,
                            ConversionRate = tbParities.ConversionRate
                        }).ToList();
            if (!string.IsNullOrEmpty(Currency))
            {
                list = list.Where(s => s.Currency.Contains(Currency.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(BasicCurrency))
            {
                list = list.Where(s => s.BasicCurrency.Contains(BasicCurrency.Trim())).ToList();
            }
            //查询数据
            List<Parities> listExaminee = list.ToList();
            //二：代码创建一个Excel表格（这里称为工作簿）
            //创建Excel文件的对象 工作簿(调用NPOI文件)
            HSSFWorkbook excelBook = new HSSFWorkbook();
            //创建Excel工作表 Sheet
            ISheet sheet1 = excelBook.CreateSheet("系统汇率");
            //给Sheet添加第一行的头部标题
            IRow row1 = sheet1.CreateRow(0);
            //给标题的每一个单元格赋值
            row1.CreateCell(0).SetCellValue("交易币");
            row1.CreateCell(1).SetCellValue("基本币");
            row1.CreateCell(2).SetCellValue("兑换率");
            //添加数据行：将表格数据逐步写入sheet1各个行中（也就是给每一个单元格赋值）
            for (int i = 0; i < listExaminee.Count; i++)
            {
                //创建行
                IRow rowTemp = sheet1.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listExaminee[i].Currency);
                rowTemp.CreateCell(1).SetCellValue(listExaminee[i].BasicCurrency);
                rowTemp.CreateCell(2).SetCellValue(listExaminee[i].ConversionRate.ToString());
            }
            //输出的文件名称
            string fileName = "系统汇率" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
            //把Excel转为流，输出
            //创建文件流
            System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
            //将工作薄写入文件流
            excelBook.Write(bookStream);
            //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
            bookStream.Seek(0, System.IO.SeekOrigin.Begin);
            //Stream对象,文件类型,文件名称
            return File(bookStream, "application/vnd.ms-excel", fileName);
        }
        /// <summary>
        /// 系统汇率报表
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievementParities(string Currency, string BasicCurrency)
        {
            var list = (from tbParities in myModels.SYS_Parities
                        select new Parities
                        {
                            ParitiesID = tbParities.ParitiesID,
                            Currency = tbParities.Currency,
                            BasicCurrency = tbParities.BasicCurrency,
                            ConversionRate = tbParities.ConversionRate
                        }).ToList();
            if (!string.IsNullOrEmpty(Currency))
            {
                list = list.Where(s => s.Currency.Contains(Currency.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(BasicCurrency))
            {
                list = list.Where(s => s.BasicCurrency.Contains(BasicCurrency.Trim())).ToList();
            }
            //把linq类型的数据listResult转化为DataTable类型数据
            DataTable dt = LINQToDataTableParities(list);
            //第一步：实例化数据集`
            PrintReport.ReportDB dbReport = new PrintReport.ReportDB();
            //第二步：将dt的数据放入数据集的数据表中
            dbReport.Tables["Parities"].Merge(dt);
            //第三步：实例化报表模板
            PrintReport.parities rp = new PrintReport.parities();
            //第四步：获取报表物理文件地址     
            string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                + "Areas\\Basics\\PrintReport\\parities.rpt";
            //第五步：把报表文件加载到ReportDocument
            rp.Load(strRptPath);
            //第六步：设置报表数据源
            rp.SetDataSource(dbReport);
            //第七步：把ReportDocument转化为文件流
            Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
            return File(stream, "application/pdf");
        }
        // //将IEnumerable<T>类型的集合转换为DataTable类型
        public DataTable LINQToDataTableParities<T>(IEnumerable<T> varlist)
        {
            //定义要返回的DataTable对象
            DataTable dtReturn = new DataTable();
            //保存列集合的属性信息数组
            PropertyInfo[] oProps = null;
            if (varlist == null) return dtReturn;//安全性检查
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
                        //得到属性的类型
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
        #region  港口资料
        /// <summary>
        /// 查询港口资料
        /// </summary>
        /// <returns></returns>
        public ActionResult selectPort(BsgridPage bsgridPage, string PortCode, string ChineseName, string EnglishName)
        {
            var list = (from tbPort in myModels.SYS_Port
                        select new port
                        {
                            PortID = tbPort.PortID,
                            PortCode = tbPort.PortCode,
                            ChineseName = tbPort.ChineseName,
                            EnglishName = tbPort.EnglishName,
                            CountryCode = tbPort.CountryCode,
                            Area = tbPort.Area,
                            Berthage = tbPort.Berthage,
                            Phone = tbPort.Phone,
                            Remark = tbPort.Remark,
                            WhetherStart = tbPort.WhetherStart
                        }).ToList();
            if (!string.IsNullOrEmpty(PortCode))
            {
                list = list.Where(s => s.PortCode.Contains(PortCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ChineseName))
            {
                list = list.Where(s => s.ChineseName.Contains(ChineseName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(EnglishName))
            {
                list = list.Where(s => s.EnglishName.Contains(EnglishName.Trim())).ToList();
            }
            int totalRow = list.Count();
            List<port> notices = list.OrderByDescending(p => p.PortID).
              Skip(bsgridPage.GetStartIndex()).
              Take(bsgridPage.pageSize).
              ToList();
            Bsgrid<port> bsgrid = new Bsgrid<port>();
            bsgrid.success = true;
            bsgrid.totalRows = totalRow;
            bsgrid.curPage = bsgridPage.curPage;
            bsgrid.data = notices;
            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 新增港口资料
        /// </summary>
        /// <returns></returns>
        public ActionResult InsertPort(SYS_Port tbport)
        {
            string strMsg = "fail";
            try
            {
                int oldCode = (from tbPort in myModels.SYS_Port
                               where tbPort.PortCode == tbport.PortCode
                               select tbPort).Count();
                if (oldCode == 0)
                {
                    int oldName = (from tbPort in myModels.SYS_Port
                                   where tbPort.ChineseName == tbport.ChineseName
                                   select tbPort).Count();
                    if (oldName == 0)
                    {
                        myModels.SYS_Port.Add(tbport);
                        if (myModels.SaveChanges() > 0)
                        {
                            strMsg = "数据新增成功！";
                        }
                    }
                    else
                    {
                        strMsg = "已有中文名！";
                    }
                }
                else
                {
                    strMsg = "已有港口代码！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 修改港口资料
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdatePort(SYS_Port pwPort)
        {
            string strMsg = "fail";
            try
            {
                if (myModels.SYS_Port.Where(m => m.PortCode == pwPort.PortCode && m.PortID != pwPort.PortID).Count() == 0)
                {
                    if (myModels.SYS_Port.Where(m => m.ChineseName == pwPort.ChineseName && m.PortID != pwPort.PortID).Count() == 0)
                    {
                        SYS_Port dbPort = (from tbport in myModels.SYS_Port
                                           where tbport.PortID == pwPort.PortID
                                           select tbport).Single();
                        dbPort.PortCode = pwPort.PortCode;
                        dbPort.ChineseName = pwPort.ChineseName;
                        dbPort.EnglishName = pwPort.EnglishName;
                        dbPort.CountryCode = pwPort.CountryCode;
                        dbPort.Area = pwPort.Area;
                        dbPort.Berthage = pwPort.Berthage;
                        dbPort.Phone = pwPort.Phone;
                        dbPort.Remark = pwPort.Remark;
                        dbPort.WhetherStart = pwPort.WhetherStart;
                        myModels.Entry(dbPort).State = System.Data.Entity.EntityState.Modified;
                        if (myModels.SaveChanges() > 0)
                        {
                            strMsg = "修改成功！";
                        }
                    }
                    else
                    {
                        strMsg = "已有中文名称，无法修改！";
                    }
                }
                else
                {
                    strMsg = "已有港口代码，无法修改！";
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 修改回填港口
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectPortById(int PortID)
        {
            try
            {
                var list = (from tbPort in myModels.SYS_Port
                            where tbPort.PortID == PortID
                            select new port
                            {
                                PortID = tbPort.PortID,
                                PortCode = tbPort.PortCode,
                                ChineseName = tbPort.ChineseName,
                                EnglishName = tbPort.EnglishName,
                                CountryCode = tbPort.CountryCode,
                                Area = tbPort.Area,
                                Berthage = tbPort.Berthage,
                                Phone = tbPort.Phone,
                                Remark = tbPort.Remark,
                                WhetherStart = tbPort.WhetherStart
                            }).Single();
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return null;
            }
        }
        /// <summary>
        /// 删除港口资料
        /// </summary>
        /// <returns></returns>
        public ActionResult DelectPort(int PortID)
        {
            string strMsg = "fail";
            try
            {
                var listPort = myModels.SYS_Port.Where(m => m.PortID == PortID).Single();
                myModels.SYS_Port.Remove(listPort);
                if (myModels.SaveChanges() > 0)
                {
                    strMsg = "删除成功！";
                }
            }
            catch (Exception e)
            {
                strMsg = "fail";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 导出港口资料
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportToExcelPort(string PortCode,string ChineseName,string EnglishName)
        {
            var list = (from tbPort in myModels.SYS_Port
                        select new port
                        {
                            PortID = tbPort.PortID,
                            PortCode = tbPort.PortCode,
                            ChineseName = tbPort.ChineseName,
                            EnglishName = tbPort.EnglishName,
                            CountryCode = tbPort.CountryCode,
                            Area = tbPort.Area,
                            Berthage = tbPort.Berthage,
                            Phone = tbPort.Phone,
                            Remark = tbPort.Remark,
                            WhetherStarta = "",
                            WhetherStartl = tbPort.WhetherStart,
                        }).ToList();
            if (!string.IsNullOrEmpty(PortCode))
            {
                list = list.Where(s => s.PortCode.Contains(PortCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ChineseName))
            {
                list = list.Where(s => s.ChineseName.Contains(ChineseName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(EnglishName))
            {
                list = list.Where(s => s.EnglishName.Contains(EnglishName.Trim())).ToList();
            }
            for(int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherStartl == true)
                {
                    list[i].WhetherStarta = "启用";
                }
                else if(list[i].WhetherStartl==false)
                {
                    list[i].WhetherStarta = "不启用";
                }
            }
            //查询数据
            List<port> listExaminee = list.ToList();
            //二：代码创建一个Excel表格（这里称为工作簿）
            //创建Excel文件的对象 工作簿(调用NPOI文件)
            HSSFWorkbook excelBook = new HSSFWorkbook();
            //创建Excel工作表 Sheet
            ISheet sheet1 = excelBook.CreateSheet("港口资料");
            //给Sheet添加第一行的头部标题
            IRow row1 = sheet1.CreateRow(0);
            //给标题的每一个单元格赋值
            row1.CreateCell(0).SetCellValue("港口代码");
            row1.CreateCell(1).SetCellValue("中文名称");
            row1.CreateCell(2).SetCellValue("英文名称");
            row1.CreateCell(3).SetCellValue("国家代码");
            row1.CreateCell(4).SetCellValue("区域");
            row1.CreateCell(5).SetCellValue("泊位数");
            row1.CreateCell(6).SetCellValue("联系电话");
            row1.CreateCell(7).SetCellValue("备注");
            row1.CreateCell(8).SetCellValue("是否启用");
            //添加数据行：将表格数据逐步写入sheet1各个行中（也就是给每一个单元格赋值）
            for (int i = 0; i < listExaminee.Count; i++)
            {
                //创建行
                IRow rowTemp = sheet1.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listExaminee[i].PortCode);
                rowTemp.CreateCell(1).SetCellValue(listExaminee[i].ChineseName);
                rowTemp.CreateCell(2).SetCellValue(listExaminee[i].EnglishName);
                rowTemp.CreateCell(3).SetCellValue(listExaminee[i].CountryCode);
                rowTemp.CreateCell(4).SetCellValue(listExaminee[i].Area);
                rowTemp.CreateCell(5).SetCellValue(listExaminee[i].Berthage.ToString());
                rowTemp.CreateCell(6).SetCellValue(listExaminee[i].Phone);
                rowTemp.CreateCell(7).SetCellValue(listExaminee[i].Remark);
                rowTemp.CreateCell(8).SetCellValue(listExaminee[i].WhetherStarta);
            }
            //输出的文件名称
            string fileName = "港口资料" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".xls";
            //把Excel转为流，输出
            //创建文件流
            System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
            //将工作薄写入文件流
            excelBook.Write(bookStream);
            //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
            bookStream.Seek(0, System.IO.SeekOrigin.Begin);
            //Stream对象,文件类型,文件名称
            return File(bookStream, "application/vnd.ms-excel", fileName);
        }
        /// <summary>
        /// 港口水晶报表
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievementPort(string PortCode, string ChineseName, string EnglishName)
        {
            var list = (from tbPort in myModels.SYS_Port
                        select new port
                        {
                            PortID = tbPort.PortID,
                            PortCode = tbPort.PortCode,
                            ChineseName = tbPort.ChineseName,
                            EnglishName = tbPort.EnglishName,
                            CountryCode = tbPort.CountryCode,
                            Area = tbPort.Area,
                            Berthage = tbPort.Berthage,
                            Phone = tbPort.Phone,
                            Remark = tbPort.Remark,
                            WhetherStarta = "",
                            WhetherStartl = tbPort.WhetherStart
                        }).ToList();
            if (!string.IsNullOrEmpty(PortCode))
            {
                list = list.Where(s => s.PortCode.Contains(PortCode.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(ChineseName))
            {
                list = list.Where(s => s.ChineseName.Contains(ChineseName.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(EnglishName))
            {
                list = list.Where(s => s.EnglishName.Contains(EnglishName.Trim())).ToList();
            }
            for (int i = 0; i < list.Count(); i++)
            {
                if (list[i].WhetherStartl == true)
                {
                    list[i].WhetherStarta = "√";
                }
                else if (list[i].WhetherStartl == false)
                {
                    list[i].WhetherStarta = "×";
                }
            }
            //把linq类型的数据listResult转化为DataTable类型数据
            DataTable dt = LINQToDataTablePort(list);
            //第一步：实例化数据集`
            PrintReport.ReportDB dbReport = new PrintReport.ReportDB();
            //第二步：将dt的数据放入数据集的数据表中
            dbReport.Tables["Port"].Merge(dt);
            //第三步：实例化报表模板
            PrintReport.port rp = new PrintReport.port();
            //第四步：获取报表物理文件地址     
            string strRptPath = System.Web.HttpContext.Current.Server.MapPath("~/")
                + "Areas\\Basics\\PrintReport\\port.rpt";
            //第五步：把报表文件加载到ReportDocument
            rp.Load(strRptPath);
            //第六步：设置报表数据源
            rp.SetDataSource(dbReport);
            //第七步：把ReportDocument转化为文件流
            Stream stream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
            return File(stream, "application/pdf");
        }
        // //将IEnumerable<T>类型的集合转换为DataTable类型
        public DataTable LINQToDataTablePort<T>(IEnumerable<T> varlist)
        {
            //定义要返回的DataTable对象
            DataTable dtReturn = new DataTable();
            //保存列集合的属性信息数组
            PropertyInfo[] oProps = null;
            if (varlist == null) return dtReturn;//安全性检查
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
                        //得到属性的类型
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
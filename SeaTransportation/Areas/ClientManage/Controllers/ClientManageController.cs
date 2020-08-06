using Hotel.VO;
using SeaTransportation.Models;
using SeaTransportation.VO;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Web.Mvc;

namespace SeaTransportation.Areas.ClientManage.Controllers
{
    public class ClientManageController : Controller
    {
        Models.SeaTransportationEntities myModels = new Models.SeaTransportationEntities();
        // GET: ClientManage/ClientManage


        #region 页面搭建

        /// <summary>
        /// 客户类型页面 
        /// </summary>
        /// <returns></returns>
        public ActionResult ClientManage()
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
            //使用ViewBag将数据返回 文本框点开回填登陆者 后面是从登陆控制器来的
            ViewBag.AccountName = Session["StaffName"].ToString().Trim();
            return View();
        }


        /// <summary>
        /// 客户标准运费
        /// </summary>
        /// <returns></returns>
        public ActionResult ClientMeasure()
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
            //使用ViewBag将数据返回 文本框点开回填登陆者 后面是从登陆控制器来的
            ViewBag.AccountName = Session["StaffName"].ToString().Trim();
            return View();
        }


        /// <summary>
        /// 客户应收运费
        /// </summary>
        /// <returns></returns>
        public ActionResult ClientReceivable()
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
            //使用ViewBag将数据返回 文本框点开回填登陆者 后面是从登陆控制器来的
            ViewBag.AccountName = Session["StaffName"].ToString().Trim();

            return View();
        }

        /// <summary>
        /// 车队标准运费
        /// </summary>
        /// <returns></returns>
        public ActionResult MotorcadeMeasure()
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
            //使用ViewBag将数据返回 文本框点开回填登陆者 后面是从登陆控制器来的
            ViewBag.AccountName = Session["StaffName"].ToString().Trim();

            return View();
        }

        /// <summary>
        /// 司机产值
        /// </summary>
        /// <returns></returns>
        public ActionResult ChauffeurProduce()
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
            //使用ViewBag将数据返回 文本框点开回填登陆者 后面是从登陆控制器来的
            ViewBag.AccountName = Session["StaffName"].ToString().Trim();

            return View();
        }


        #endregion

        #region 查询表格数据

        /// <summary>
        /// 查询客户关系管理表格数据
        /// </summary>
        /// <param name="bsgridPage"></param>
        /// <returns></returns>
        public ActionResult SelectClient(BsgridPage bsgridPage,int? ClientID,string AAA, string ClientCode,string ClientAbbreviation,int? GatedotID,int? MessageID,bool? WhetherStart,string ClientType)
        {
            var listClient = (from tbClient in myModels.SYS_Client
                              join tbGatedot in myModels.SYS_Gatedot on tbClient.GatedotID equals tbGatedot.GatedotID
                              join tbstaff in myModels.SYS_Staff on tbClient.StaffID equals tbstaff.StaffID
                              orderby tbClient.ClientID descending
                              select new ClientVo
                              {
                                  ClientID = tbClient.ClientID,//客户信息ID
                                  ClientCode = tbClient.ClientCode,//客户代码
                                  ClientAbbreviation = tbClient.ClientAbbreviation,//客户简称
                                  ChineseName = tbClient.ChineseName,//中文名称
                                  ClientRank = tbClient.ClientRank,//客户级别
                                  CustomsCode = tbClient.CustomsCode,//海关编码
                                  GatedotName = tbGatedot.GatedotName,//所属区域
                                  GatedotID = tbGatedot.GatedotID,//所属区域ID（门点ID）
                                  ClientSource = tbClient.ClientSource,//客户来源
                                  ClientPhone = tbClient.ClientPhone,//客户电话
                                  WhetherStart = tbClient.WhetherStart,//是否启用
                                  WhetherStartt = "",
                                  MessageID = tbClient.MessageID,//组织结构ID
                                  StaffName = tbstaff.StaffName,// 业务员
                                  ClientFax = tbClient.ClientFax,//客户传真
                                  Email = tbClient.Email,
                                  PostCode = tbClient.PostCode,//邮编
                                  OfficeHours1 = tbClient.OfficeHours.ToString(),//上班时间
                                  ClosingTime1 = tbClient.ClosingTime.ToString(),//下班时间
                                  Site = tbClient.Site,//地址
                                  Website = tbClient.Website,//网站
                                  OpenAccount = tbClient.OpenAccount,//开户行
                                  OpenAccountCode = tbClient.OpenAccountCode,//开户行账户
                                  Describe = tbClient.Describe,//描述
                              }).ToList();

            for (int i = 0; i < listClient.Count(); i++)  //转换true false 为是否
            {
                if (listClient[i].WhetherStart ==true)
                {
                    listClient[i].WhetherStartt = "是";
                }
                else
                {
                    listClient[i].WhetherStartt = "否";
                }
            }

            if (AAA == "AAA")//页面传过来的参数
            {
                listClient = listClient.Where(m => m.ClientID == ClientID).ToList();
                return Json(listClient, JsonRequestBehavior.AllowGet);
            }
            if (!string.IsNullOrEmpty(ClientCode))//客户代码
            {
                ClientCode = ClientCode.Trim();
                listClient = listClient.Where(p => p.ClientCode.Contains(ClientCode)).ToList();
            }
            if (!string.IsNullOrEmpty(ClientAbbreviation))//客户简称
            {
                ClientAbbreviation = ClientAbbreviation.Trim();
                listClient = listClient.Where(p => p.ClientAbbreviation.Contains(ClientAbbreviation)).ToList();
            }
            if (GatedotID > 0)//所属区域ID（门点ID）
            {
                listClient = listClient.Where(m => m.GatedotID == GatedotID).ToList();
            }
            if (MessageID > 0)//组织结构ID
            {
                listClient = listClient.Where(m => m.MessageID == MessageID).ToList();
            }
            if (WhetherStart!=null)//是否启用  只要判断不为空就可以了
            {
                listClient = listClient.Where(m => m.WhetherStart == WhetherStart).ToList();
            }
            if (!string.IsNullOrEmpty(ClientType)) //客户类型
            {
                if (listClient.Count()>0)
                {
                    for (int i = 0; i < listClient.Count(); i++)
                    {
                        var clientID = listClient[i].ClientID;
                        if (myModels.SYS_ClientType.Where(m=>m.ClientID== clientID&&m.ClientType.Trim()== ClientType.Trim()).Count()==0)
                        {
                            listClient.Remove(listClient[i]);
                        }
                    }
                }
            }

            //获取当前查询出来的条数
            var totalCount = listClient.Count();

            List<ClientVo> listItem = listClient
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ClientVo> bsgrid = new Bsgrid<ClientVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// 查询联系人表格
        /// </summary>
        /// <param name="bsgridPage"></param>
        /// <returns></returns>
        public ActionResult SelectScontacts(BsgridPage bsgridPage,int? ClientID,int? ClientScontactsID,string BBB)  //参数要有值
        {
            var listScontacts = (from tbScontacts in myModels.SYS_ClientScontacts
                                where tbScontacts.ClientID == ClientID
                                orderby tbScontacts.ClientID descending
                                select new ClientVo
                                {
                                    ClientID = tbScontacts.ClientID,
                                    ClientScontactsID = tbScontacts.ClientScontactsID,
                                    contactsPhone = tbScontacts.contactsPhone,
                                    contactsName = tbScontacts.contactsName,
                                }).ToList();
            if (BBB == "BBB")//页面传过来的参数
            {
                listScontacts = listScontacts.Where(m => m.ClientScontactsID == ClientScontactsID).ToList();
                return Json(listScontacts, JsonRequestBehavior.AllowGet);
            }

            //获取当前查询出来的条数
            var totalCount = listScontacts.Count();

            List<ClientVo> listItem = listScontacts
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ClientVo> bsgrid = new Bsgrid<ClientVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询工厂地址表格数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectClientSite(BsgridPage bsgridPage,int? ClientSiteID, string BBB, int? ClientID)
        {
            var listClientSite = (from tbClientSite in myModels.SYS_ClientSite
                                  join tbGatedot in myModels.SYS_Gatedot on tbClientSite.GatedotID equals tbGatedot.GatedotID
                                  where tbClientSite.ClientID == ClientID
                                  orderby tbClientSite.ClientSiteID descending
                                  select new ClientVo
                                  {
                                      ClientSiteID = tbClientSite.ClientSiteID,
                                      ClientID = tbClientSite.ClientID,
                                      ContactsName = tbClientSite.ContactsName,
                                      ContactsPhone = tbClientSite.ContactsPhone,
                                      ClientSite = tbClientSite.ClientSite,
                                      FactoryName = tbClientSite.FactoryName,
                                      FactoryCode = tbClientSite.FactoryCode,
                                      GatedotName = tbGatedot.GatedotName,
                                      Remark = tbClientSite.Remark,
                                      GatedotID = tbClientSite.GatedotID
                                  }).ToList();

            if (BBB == "BBB")//页面传过来的参数
            {
                listClientSite = listClientSite.Where(m => m.ClientSiteID == ClientSiteID).ToList();
                return Json(listClientSite, JsonRequestBehavior.AllowGet);
            }

            //获取当前查询出来的条数
            var totalCount = listClientSite.Count();

            List<ClientVo> listItem = listClientSite
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ClientVo> bsgrid = new Bsgrid<ClientVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// 查询客户标准运费上方表格数据的报价表
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectOffer(BsgridPage bsgridPage)
        {
            var listOffer = (from tbOffer in myModels.SYS_Offer
                             where tbOffer.OfferType.Trim()== "客户标准运费"
                             orderby tbOffer.OfferID descending
                             select new ClientVo
                             {
                                 OfferID = tbOffer.OfferID,
                                 OfferDate1 = tbOffer.OfferDate.ToString(),
                                 TakeEffectDate1 = tbOffer.TakeEffectDate.ToString(),
                                 LoseEfficacyDate1 = tbOffer.LoseEfficacyDate.ToString(),
                                 WhetherShuii = "",
                                 WhetherShui = tbOffer.WhetherShui,
                                 Remark = tbOffer.Remark,
                             }).ToList();
            for (int i = 0; i < listOffer.Count(); i++)  //转换true false 为是否
            {
                if (listOffer[i].WhetherShui == true)
                {
                    listOffer[i].WhetherShuii = "是";
                }
                else
                {
                    listOffer[i].WhetherShuii = "否";
                }
            }

            //获取当前查询出来的条数
            var totalCount = listOffer.Count();

            List<ClientVo> listItem = listOffer
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ClientVo> bsgrid = new Bsgrid<ClientVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询客户标准运费下方表格数据的报价明细
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectOfferDetail(BsgridPage bsgridPage,int? OfferID,int? OfferDetailID,string AAA)
        {
            var listOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                   join tbExpense in myModels.SYS_Expense on tbOfferDetail.ExpenseID equals tbExpense.ExpenseID
                                   join tbHaulWay in myModels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWay.HaulWayID
                                   join tbMention in myModels.SYS_Mention on tbHaulWay.MentionAreaID equals tbMention.MentionID
                                   join tbLoading in myModels.SYS_Gatedot on tbHaulWay.LoadingAreaID equals tbLoading.GatedotID
                                   join tbAlsoTank in myModels.SYS_Mention on tbHaulWay.AlsoTankAreaID equals tbAlsoTank.MentionID
                                   join tbStaff in myModels.SYS_Staff on tbOfferDetail.StaffID equals tbStaff.StaffID
                                   orderby tbOfferDetail.OfferDetailID descending
                                   select new ClientVo
                                   {
                                       HaulWayID = tbHaulWay.HaulWayID,
                                       OfferID = tbOfferDetail.OfferID,
                                       OfferDetailID = tbOfferDetail.OfferDetailID,
                                       Abbreviation = tbMention.Abbreviation,
                                       Abbreviationn = tbAlsoTank.Abbreviation,
                                       GatedotName = tbLoading.GatedotName,
                                       HaulWayDescription = tbMention.Abbreviation + "-" + tbLoading.GatedotName + "-" + tbAlsoTank.Abbreviation,
                                       ExpenseName = tbExpense.ExpenseName,
                                       CabinetType = tbOfferDetail.CabinetType.Trim(),
                                       EtryClasses = tbOfferDetail.EtryClasses.Trim(),
                                       Money = tbOfferDetail.Money,
                                       Currency = tbOfferDetail.Currency,
                                       Remark = tbOfferDetail.Remark,
                                       StaffName = tbStaff.StaffName,
                                       BoxQuantity = tbOfferDetail.BoxQuantity
                                   }).ToList();

            if (OfferID > 0)
            {
                listOfferDetail = listOfferDetail.Where(m => m.OfferID == OfferID).ToList();
            }
            if (AAA == "AAA")//页面传过来的参数
            {
                listOfferDetail = listOfferDetail.Where(m => m.OfferDetailID == OfferDetailID).ToList();
                return Json(listOfferDetail, JsonRequestBehavior.AllowGet);
            }

            //获取当前查询出来的条数
            var totalCount = listOfferDetail.Count();

            List<ClientVo> listItem = listOfferDetail
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ClientVo> bsgrid = new Bsgrid<ClientVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询客户应收运费上方表格报价数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectReceivableOffer(BsgridPage bsgridPage,int? OfferID,string ClientAbbreviation,string OfferDate,string OfferDatee,bool? WhetherShui,string AAA,int? ClientID)
        {
            var listOffer = (from tbOffer in myModels.SYS_Offer
                             join tbClient in myModels.SYS_Client on tbOffer.ClientID equals tbClient.ClientID
                             where tbOffer.OfferType.Trim() == "客户应收费用"
                             orderby tbOffer.OfferID descending
                             select new ClientVo
                             {
                                 OfferID = tbOffer.OfferID,
                                 OfferDate1 = tbOffer.OfferDate.ToString(),
                                 OfferDate = tbOffer.OfferDate,
                                 TakeEffectDate1 = tbOffer.TakeEffectDate.ToString(),
                                 LoseEfficacyDate1 = tbOffer.LoseEfficacyDate.ToString(),
                                 WhetherShuii = "",
                                 WhetherShui = tbOffer.WhetherShui,
                                 Remark = tbOffer.Remark,
                                 ClientAbbreviation = tbClient.ClientAbbreviation,
                                 ClientCode = tbClient.ClientCode,
                                 ClientID = tbClient.ClientID,

                             }).ToList();
            for (int i = 0; i < listOffer.Count(); i++)  //转换true false 为是否
            {
                if (listOffer[i].WhetherShui == true)
                {
                    listOffer[i].WhetherShuii = "是";
                }
                else
                {
                    listOffer[i].WhetherShuii = "否";
                }
            }

            if (!string.IsNullOrEmpty(ClientAbbreviation))//客户简称
            {
                ClientAbbreviation = ClientAbbreviation.Trim();
                listOffer = listOffer.Where(p => p.ClientAbbreviation.Contains(ClientAbbreviation)).ToList();
            }

            if (ClientID > 0)
            {
                listOffer = listOffer.Where(m => m.ClientID == ClientID).ToList();
            }

            //判断时间段去查询的方法，有三种情况，所以需要写三个判断
            if (!string.IsNullOrEmpty(OfferDate) && string.IsNullOrEmpty(OfferDatee))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(OfferDate);
                    listOffer = listOffer.Where(m => m.OfferDate == Time).ToList();
                }
                catch (Exception)
                {
                    listOffer = listOffer.Where(m => m.OfferID == OfferID).ToList();
                }
            }
            else if (string.IsNullOrEmpty(OfferDate) && !string.IsNullOrEmpty(OfferDatee))
            {
                try
                {
                    DateTime Times = Convert.ToDateTime(OfferDatee);
                    listOffer = listOffer.Where(m => m.OfferDate <= Times).ToList();
                }
                catch (Exception)
                {
                    listOffer = listOffer.Where(m => m.OfferID == OfferID).ToList();
                }
            }
            else if (!string.IsNullOrEmpty(OfferDate) && !string.IsNullOrEmpty(OfferDatee))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(OfferDate);
                    DateTime Times = Convert.ToDateTime(OfferDatee);
                    listOffer = listOffer.Where(m => m.OfferDate >= Time && m.OfferDate <= Times).ToList();
                }
                catch (Exception)
                {
                    listOffer = listOffer.Where(m => m.OfferID == OfferID).ToList();
                }
            }

            if (WhetherShui != null)//是否含税  只要判断不为空就可以了
            {
                listOffer = listOffer.Where(m => m.WhetherShui == WhetherShui).ToList();
            }

            if (AAA == "AAA")//页面传过来的参数
            {
                listOffer = listOffer.Where(m => m.OfferID == OfferID).ToList();
                return Json(listOffer, JsonRequestBehavior.AllowGet);
            }

            //获取当前查询出来的条数
            var totalCount = listOffer.Count();

            List<ClientVo> listItem = listOffer
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ClientVo> bsgrid = new Bsgrid<ClientVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询客户应收运费下方表格报价明细数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectReceivableOfferDetail(BsgridPage bsgridPage, int? OfferID,int? OfferDetailID,string AAA)
        {
            var listOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                   join tbExpense in myModels.SYS_Expense on tbOfferDetail.ExpenseID equals tbExpense.ExpenseID
                                   join tbHaulWay in myModels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWay.HaulWayID
                                   join tbMention in myModels.SYS_Mention on tbHaulWay.MentionAreaID equals tbMention.MentionID
                                   join tbLoading in myModels.SYS_Gatedot on tbHaulWay.LoadingAreaID equals tbLoading.GatedotID
                                   join tbAlsoTank in myModels.SYS_Mention on tbHaulWay.AlsoTankAreaID equals tbAlsoTank.MentionID
                                   join tbStaff in myModels.SYS_Staff on tbOfferDetail.StaffID equals tbStaff.StaffID
                                   orderby tbOfferDetail.OfferDetailID descending
                                   select new ClientVo
                                   {
                                       HaulWayID = tbHaulWay.HaulWayID,
                                       OfferID = tbOfferDetail.OfferID,
                                       OfferDetailID = tbOfferDetail.OfferDetailID,
                                       Abbreviation = tbMention.Abbreviation.Trim(),
                                       Abbreviationn = tbAlsoTank.Abbreviation.Trim(),
                                       GatedotName = tbLoading.GatedotName.Trim(),
                                       HaulWayDescription = tbMention.Abbreviation.Trim() + "-" + tbLoading.GatedotName.Trim() + "-" + tbAlsoTank.Abbreviation.Trim(),
                                       ExpenseName = tbExpense.ExpenseName,
                                       CabinetType = tbOfferDetail.CabinetType.Trim(),
                                       EtryClasses = tbOfferDetail.EtryClasses.Trim(),
                                       Money = tbOfferDetail.Money,
                                       Currency = tbOfferDetail.Currency,
                                       Remark = tbOfferDetail.Remark,
                                       StaffName = tbStaff.StaffName,
                                       BoxQuantity = tbOfferDetail.BoxQuantity,
                                   }).ToList();

            if (OfferID > 0)
            {
                listOfferDetail = listOfferDetail.Where(m => m.OfferID == OfferID).ToList();
            }

            if (AAA == "AAA")//页面传过来的参数
            {
                listOfferDetail = listOfferDetail.Where(m => m.OfferDetailID == OfferDetailID).ToList();
                return Json(listOfferDetail, JsonRequestBehavior.AllowGet);
            }

            //获取当前查询出来的条数
            var totalCount = listOfferDetail.Count();

            List<ClientVo> listItem = listOfferDetail
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ClientVo> bsgrid = new Bsgrid<ClientVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 选择下拉框的数据，回填文本框的值的在数据库查询的方法
        /// </summary>
        /// <param name="ClientID"></param>
        /// <returns></returns>
        public ActionResult SelectCClient(int? ClientID)
        {
            if (ClientID>=1)
            {
                var listClient = myModels.SYS_Client.Where(m => m.ClientID == ClientID).Select(m => m.ClientAbbreviation.Trim()).ToList();
                return Json(listClient[0], JsonRequestBehavior.AllowGet);
            }
            else
            {
                return null;
            }
        }


        /// <summary>
        /// 查询车队标准运费的上方表格报价数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectMotorcadeOffce(BsgridPage bsgridPage,int? ClientID, int? OfferID, string ClientAbbreviation, string OfferDate, string OfferDatee, bool? WhetherShui, string AAA)
        {
            var listOffer = (from tbOffer in myModels.SYS_Offer
                             join tbClient in myModels.SYS_Client on tbOffer.ClientID equals tbClient.ClientID
                             where tbOffer.OfferType.Trim() == "车队标准运费"
                             orderby tbOffer.OfferID descending
                             select new ClientVo
                             {
                                 OfferID = tbOffer.OfferID,
                                 OfferDate1 = tbOffer.OfferDate.ToString(),
                                 OfferDate = tbOffer.OfferDate,
                                 TakeEffectDate1 = tbOffer.TakeEffectDate.ToString(),
                                 LoseEfficacyDate1 = tbOffer.LoseEfficacyDate.ToString(),
                                 WhetherShuii = "",
                                 WhetherShui = tbOffer.WhetherShui,
                                 Remark = tbOffer.Remark,
                                 ClientAbbreviation = tbClient.ClientAbbreviation,
                                 ClientCode = tbClient.ClientCode,
                                 ClientID = tbClient.ClientID,

                             }).ToList();

            for (int i = 0; i < listOffer.Count(); i++)  //转换true false 为是否
            {
                if (listOffer[i].WhetherShui == true)
                {
                    listOffer[i].WhetherShuii = "是";
                }
                else
                {
                    listOffer[i].WhetherShuii = "否";
                }
            }

            if (ClientID > 0)
            {
                listOffer = listOffer.Where(m => m.ClientID == ClientID).ToList();
            }
            if (!string.IsNullOrEmpty(ClientAbbreviation))//客户简称
            {
                ClientAbbreviation = ClientAbbreviation.Trim();
                listOffer = listOffer.Where(p => p.ClientAbbreviation.Contains(ClientAbbreviation)).ToList();
            }

            //判断时间段去查询的方法，有三种情况，所以需要写三个判断
            if (!string.IsNullOrEmpty(OfferDate) && string.IsNullOrEmpty(OfferDatee))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(OfferDate);
                    listOffer = listOffer.Where(m => m.OfferDate == Time).ToList();
                }
                catch (Exception)
                {
                    listOffer = listOffer.Where(m => m.OfferID == OfferID).ToList();
                }
            }
            else if (string.IsNullOrEmpty(OfferDate) && !string.IsNullOrEmpty(OfferDatee))
            {
                try
                {
                    DateTime Times = Convert.ToDateTime(OfferDatee);
                    listOffer = listOffer.Where(m => m.OfferDate <= Times).ToList();
                }
                catch (Exception)
                {
                    listOffer = listOffer.Where(m => m.OfferID == OfferID).ToList();
                }
            }
            else if (!string.IsNullOrEmpty(OfferDate) && !string.IsNullOrEmpty(OfferDatee))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(OfferDate);
                    DateTime Times = Convert.ToDateTime(OfferDatee);
                    listOffer = listOffer.Where(m => m.OfferDate >= Time && m.OfferDate <= Times).ToList();
                }
                catch (Exception)
                {
                    listOffer = listOffer.Where(m => m.OfferID == OfferID).ToList();
                }
            }

            if (WhetherShui != null)//是否含税  只要判断不为空就可以了
            {
                listOffer = listOffer.Where(m => m.WhetherShui == WhetherShui).ToList();
            }

            if (AAA == "AAA")//页面传过来的参数
            {
                listOffer = listOffer.Where(m => m.OfferID == OfferID).ToList();
                return Json(listOffer, JsonRequestBehavior.AllowGet);
            }


            //获取当前查询出来的条数
            var totalCount = listOffer.Count();

            List<ClientVo> listItem = listOffer
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ClientVo> bsgrid = new Bsgrid<ClientVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询车队标准运费的下方表格的报价明细数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectMotorcadeOfferDetail(BsgridPage bsgridPage, int? OfferID, int? OfferDetailID, string AAA)
        {
            var listOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                   join tbExpense in myModels.SYS_Expense on tbOfferDetail.ExpenseID equals tbExpense.ExpenseID
                                   join tbHaulWay in myModels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWay.HaulWayID
                                   join tbMention in myModels.SYS_Mention on tbHaulWay.MentionAreaID equals tbMention.MentionID
                                   join tbLoading in myModels.SYS_Gatedot on tbHaulWay.LoadingAreaID equals tbLoading.GatedotID
                                   join tbAlsoTank in myModels.SYS_Mention on tbHaulWay.AlsoTankAreaID equals tbAlsoTank.MentionID
                                   join tbStaff in myModels.SYS_Staff on tbOfferDetail.StaffID equals tbStaff.StaffID
                                   orderby tbOfferDetail.OfferDetailID descending
                                   select new ClientVo
                                   {
                                       HaulWayID = tbHaulWay.HaulWayID,
                                       OfferID = tbOfferDetail.OfferID,
                                       OfferDetailID = tbOfferDetail.OfferDetailID,
                                       Abbreviation = tbMention.Abbreviation.Trim(),
                                       Abbreviationn = tbAlsoTank.Abbreviation.Trim(),
                                       GatedotName = tbLoading.GatedotName.Trim(),
                                       HaulWayDescription = tbMention.Abbreviation.Trim() + "-" + tbLoading.GatedotName.Trim() + "-" + tbAlsoTank.Abbreviation.Trim(),
                                       ExpenseName = tbExpense.ExpenseName,
                                       CabinetType = tbOfferDetail.CabinetType.Trim(),
                                       EtryClasses = tbOfferDetail.EtryClasses.Trim(),
                                       Money = tbOfferDetail.Money,
                                       Currency = tbOfferDetail.Currency,
                                       Remark = tbOfferDetail.Remark,
                                       StaffName = tbStaff.StaffName,
                                       BoxQuantity = tbOfferDetail.BoxQuantity,
                                   }).ToList();

            if (OfferID > 0)
            {
                listOfferDetail = listOfferDetail.Where(m => m.OfferID == OfferID).ToList();
            }

            if (AAA == "AAA")//页面传过来的参数
            {
                listOfferDetail = listOfferDetail.Where(m => m.OfferDetailID == OfferDetailID).ToList();
                return Json(listOfferDetail, JsonRequestBehavior.AllowGet);
            }

            //获取当前查询出来的条数
            var totalCount = listOfferDetail.Count();

            List<ClientVo> listItem = listOfferDetail
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ClientVo> bsgrid = new Bsgrid<ClientVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询司机产值的上方表格报价表格的数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectChauffeurProduceOffer(BsgridPage bsgridPage)
        {
            var listChauff = (from tbChauff in myModels.SYS_Offer
                              where tbChauff.OfferType.Trim() == "司机产值"
                              orderby tbChauff.OfferID descending
                              select new ClientVo
                              {
                                  OfferID = tbChauff.OfferID,
                                  OfferDate1 = tbChauff.OfferDate.ToString(),
                                  TakeEffectDate1 = tbChauff.TakeEffectDate.ToString(),
                                  LoseEfficacyDate1 = tbChauff.LoseEfficacyDate.ToString(),
                                  WhetherShuii = "",
                                  WhetherShui = tbChauff.WhetherShui,
                                  Remark = tbChauff.Remark
                              }).ToList();
            for (int i = 0; i < listChauff.Count(); i++)  //转换true false 为是否
            {
                if (listChauff[i].WhetherShui == true)
                {
                    listChauff[i].WhetherShuii = "是";
                }
                else
                {
                    listChauff[i].WhetherShuii = "否";
                }
            }

            //获取当前查询出来的条数
            var totalCount = listChauff.Count();

            List<ClientVo> listItem = listChauff
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ClientVo> bsgrid = new Bsgrid<ClientVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询司机产值的下方表格报价明细表格的数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelecttChauffeurProduceOfferDetail(BsgridPage bsgridPage, int? OfferID, int? OfferDetailID, string AAA)
        {
            var listOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                   join tbExpense in myModels.SYS_Expense on tbOfferDetail.ExpenseID equals tbExpense.ExpenseID
                                   join tbHaulWay in myModels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWay.HaulWayID
                                   join tbMention in myModels.SYS_Mention on tbHaulWay.MentionAreaID equals tbMention.MentionID
                                   join tbLoading in myModels.SYS_Gatedot on tbHaulWay.LoadingAreaID equals tbLoading.GatedotID
                                   join tbAlsoTank in myModels.SYS_Mention on tbHaulWay.AlsoTankAreaID equals tbAlsoTank.MentionID
                                   join tbStaff in myModels.SYS_Staff on tbOfferDetail.StaffID equals tbStaff.StaffID
                                   orderby tbOfferDetail.OfferDetailID descending
                                   select new ClientVo
                                   {
                                       HaulWayID = tbHaulWay.HaulWayID,
                                       OfferID = tbOfferDetail.OfferID,
                                       OfferDetailID = tbOfferDetail.OfferDetailID,
                                       Abbreviation = tbMention.Abbreviation.Trim(),
                                       Abbreviationn = tbAlsoTank.Abbreviation.Trim(),
                                       GatedotName = tbLoading.GatedotName.Trim(),
                                       HaulWayDescription = tbMention.Abbreviation + "-" + tbLoading.GatedotName + "-" + tbAlsoTank.Abbreviation,
                                       ExpenseName = tbExpense.ExpenseName,
                                       CabinetType = tbOfferDetail.CabinetType.Trim(),
                                       EtryClasses = tbOfferDetail.EtryClasses.Trim(),
                                       Money = tbOfferDetail.Money,
                                       Currency = tbOfferDetail.Currency,
                                       Remark = tbOfferDetail.Remark,
                                       StaffName = tbStaff.StaffName,
                                       BoxQuantity = tbOfferDetail.BoxQuantity
                                   }).ToList();

            if (OfferID > 0)
            {
                listOfferDetail = listOfferDetail.Where(m => m.OfferID == OfferID).ToList();
            }
            if (AAA == "AAA")//页面传过来的参数   (用于表单数据回填时)
            {
                listOfferDetail = listOfferDetail.Where(m => m.OfferDetailID == OfferDetailID).ToList();
                return Json(listOfferDetail, JsonRequestBehavior.AllowGet);
            }

            //获取当前查询出来的条数
            var totalCount = listOfferDetail.Count();

            List<ClientVo> listItem = listOfferDetail
                                       .Skip(bsgridPage.GetStartIndex())
                                       .Take(bsgridPage.pageSize)
                                       .ToList();

            Bsgrid<ClientVo> bsgrid = new Bsgrid<ClientVo>()
            {
                success = true,
                totalRows = totalCount,
                curPage = bsgridPage.curPage,
                data = listItem
            };

            return Json(bsgrid, JsonRequestBehavior.AllowGet);
        }

        #endregion

        #region 保存代码

        /// <summary>
        /// 保存新增客户信息
        /// </summary>
        /// <returns></returns>
        public ActionResult SaveClient(SYS_Client myClient,bool? EntrustUnit,bool? Consigner,bool? Ship,bool? Partner,bool? Delivery,bool? Trailer,bool? Insurance,bool? Agency,bool? Inform,bool? Logistics,bool? Clearance)
        {
            string strMsg = "failed";

            //实例化
            SYS_Client Client = new SYS_Client(); //客户类型表

            var ClientCodee = (from tb in myModels.SYS_Client
                               where tb.ClientCode == myClient.ClientCode
                               select tb).Count();
            if (ClientCodee==0)
            {
                //写入数据
                myClient.StaffID = Convert.ToInt32(Session["StaffID"].ToString()); //强势转换

                //写入数据库
                myModels.SYS_Client.Add(myClient);
            }

            //保存数据库
            if (myModels.SaveChanges() > 0)
            {
                //实例化
                SYS_ClientType myType = new SYS_ClientType(); //客户类型表

                if (EntrustUnit==true)//委托单位
                {
                    myType.ClientID = myClient.ClientID;//客户ID
                    myType.ClientType = "委托单位";
                    myModels.SYS_ClientType.Add(myType);
                    myModels.SaveChanges();
                }
                if (Consigner==true)//发货人
                {
                    myType.ClientID = myClient.ClientID;//客户ID
                    myType.ClientType = "发货人";
                    myModels.SYS_ClientType.Add(myType);
                    myModels.SaveChanges();
                }
                if (Ship==true)//船公司
                {
                    myType.ClientID = myClient.ClientID;//客户ID
                    myType.ClientType = "船公司";
                    myModels.SYS_ClientType.Add(myType);
                    myModels.SaveChanges();
                }
                if (Partner==true)//合作代理
                {
                    myType.ClientID = myClient.ClientID;//客户ID
                    myType.ClientType = "合作代理";
                    myModels.SYS_ClientType.Add(myType);
                    myModels.SaveChanges();
                }
                if (Delivery==true)//收货人
                {
                    myType.ClientID = myClient.ClientID;//客户ID
                    myType.ClientType = "收货人";
                    myModels.SYS_ClientType.Add(myType);
                    myModels.SaveChanges();
                }
                if (Trailer==true)//拖车公司
                {
                    myType.ClientID = myClient.ClientID;//客户ID
                    myType.ClientType = "拖车公司";
                    myModels.SYS_ClientType.Add(myType);
                    myModels.SaveChanges();
                }
                if (Insurance==true)//保险公司
                {
                    myType.ClientID = myClient.ClientID;//客户ID
                    myType.ClientType = "保险公司";
                    myModels.SYS_ClientType.Add(myType);
                    myModels.SaveChanges();
                }
                if (Agency==true)//船代公司
                {
                    myType.ClientID = myClient.ClientID;//客户ID
                    myType.ClientType = "船代公司";
                    myModels.SYS_ClientType.Add(myType);
                    myModels.SaveChanges();
                }
                if (Inform==true)//通知人
                {
                    myType.ClientID = myClient.ClientID;//客户ID
                    myType.ClientType = "通知人";
                    myModels.SYS_ClientType.Add(myType);
                    myModels.SaveChanges();
                }
                if (Logistics==true)//物流公司
                {
                    myType.ClientID = myClient.ClientID;//客户ID
                    myType.ClientType = "物流公司";
                    myModels.SYS_ClientType.Add(myType);
                    myModels.SaveChanges();
                }
                if (Clearance==true)//报关行
                {
                    myType.ClientID = myClient.ClientID;//客户ID
                    myType.ClientType = "报关行";
                    myModels.SYS_ClientType.Add(myType);
                    myModels.SaveChanges();
                }

                strMsg = "success";
            }
            else
            {
                strMsg = "failed";
            }

            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// 保存修改数据
        /// </summary>
        /// <param name="myClient"></param>
        /// <returns></returns>
        public ActionResult SaveUpdate(SYS_Client myClient)
        {
            string strMsg = "faileed";
            try
            {
                SYS_Client Client = (from tbClientt in myModels.SYS_Client
                                        where tbClientt.ClientID == myClient.ClientID
                                        select tbClientt).Single();
                //修改数据
                Client.ClientID = myClient.ClientID;
                Client.ClientAbbreviation = myClient.ClientAbbreviation;
                Client.ClientCode = myClient.ClientCode;
                Client.ChineseName = myClient.ChineseName;
                Client.MessageID = myClient.MessageID;
                Client.WhetherStart = myClient.WhetherStart;
                Client.ClientSource = myClient.ClientSource;
                Client.ClientRank = myClient.ClientRank;
                Client.StaffID = Convert.ToInt32(Session["StaffID"].ToString());
                Client.CustomsCode = myClient.CustomsCode;
                Client.GatedotID = myClient.GatedotID;
                Client.ClientFax = myClient.ClientFax;
                Client.Email = myClient.Email;
                Client.OfficeHours = myClient.OfficeHours;
                Client.ClosingTime = myClient.ClosingTime;
                Client.PostCode = myClient.PostCode;
                Client.Site = myClient.Site;
                Client.Website = myClient.Website;
                Client.OpenAccount = myClient.OpenAccount;
                Client.OpenAccountCode = myClient.OpenAccountCode;
                Client.Describe = myClient.Describe;

                myModels.Entry(Client).State = System.Data.Entity.EntityState.Modified;
                myModels.SaveChanges();
                strMsg = "success";
                return Json(strMsg, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                strMsg = "faileed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// 保存删除代码
        /// </summary>
        /// <returns></returns>
        public ActionResult DeleteClient(int ClientID)
        {
            string strMsg = "failed";
            try
            {
                //实例化
                List<SYS_ClientType> listType = (from tbType in myModels.SYS_ClientType
                                                 where tbType.ClientID == ClientID
                                                 select tbType).ToList();   //如果两个表同时操作，可以传两张表同时有的ID来查出需要操作的数据

                myModels.SYS_ClientType.RemoveRange(listType);

                if (myModels.SaveChanges() > 0)
                {
                    List<SYS_Client> listClient = (from tbClient in myModels.SYS_Client
                                                   where tbClient.ClientID == ClientID
                                                   select tbClient).ToList(); //查出要删除的数据

                    myModels.SYS_Client.RemoveRange(listClient);
                    myModels.SaveChanges();
                    strMsg = "success";
                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }

            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存新增联系人代码
        /// </summary>
        /// <returns></returns>
        public ActionResult SaveClientScontacts(int ClientID,string contactsName,string contactsPhone)
        {
            string strMsg = "failed";
            try
            {
                //实例化
                SYS_ClientScontacts myClientScontacts = new SYS_ClientScontacts(); //客户类型表

                var contactsPhonee = (from tb in myModels.SYS_ClientScontacts
                                 where tb.contactsPhone == contactsPhone
                                      select tb).Count();
                if (contactsPhonee==0)
                {
                    //写入数据
                    myClientScontacts.ClientID = ClientID;
                    myClientScontacts.contactsName = contactsName;
                    myClientScontacts.contactsPhone = contactsPhone;

                    // 写入数据库
                    myModels.SYS_ClientScontacts.Add(myClientScontacts);
                    //保存数据库
                    myModels.SaveChanges();
                    strMsg = "success";
                }
                else
                {
                    strMsg = "failed";
                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存修改联系人
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateClientScontacts(int ClientScontactsID,int ClientID, string contactsName, string contactsPhone)
        {
            string strMsg = "failed";
            try
            {
                SYS_ClientScontacts myClientScontacts = (from tbClientScontacts in myModels.SYS_ClientScontacts
                                                         where tbClientScontacts.ClientScontactsID == ClientScontactsID
                                                         select tbClientScontacts).Single();
                //修改数据
                myClientScontacts.ClientScontactsID = ClientScontactsID;
                myClientScontacts.ClientID = ClientID;
                myClientScontacts.contactsName = contactsName;
                myClientScontacts.contactsPhone = contactsPhone;

                myModels.Entry(myClientScontacts).State = System.Data.Entity.EntityState.Modified;
                myModels.SaveChanges();
                strMsg = "success";
                return Json(strMsg, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存删除联系人
        /// </summary>
        /// <param name="ClientScontactsID"></param>
        /// <returns></returns>
        public ActionResult DeleteClientScontacts(int ClientScontactsID)
        {
            string strMsg = "failed";
            try
            {
                //实例化
                List<SYS_ClientScontacts> listClientScontacts = (from tbClientScontacts in myModels.SYS_ClientScontacts
                                                                 where tbClientScontacts.ClientScontactsID == ClientScontactsID
                                                                 select tbClientScontacts).ToList();   //如果两个表同时操作，可以传两张表同时有的ID来查出需要操作的数据

                myModels.SYS_ClientScontacts.RemoveRange(listClientScontacts);
                myModels.SaveChanges();
                strMsg = "success";
            }
            catch (Exception)
            {
                strMsg = "failed";
            }

            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存新增工厂地址
        /// </summary>
        /// <returns></returns>
        public ActionResult SaveClientSite(SYS_ClientSite myClientSite,int ClientID)
        {
            string strMsg = "failed";
            try
            {
                var FactoryCodee = (from tb in myModels.SYS_ClientSite
                                    where tb.FactoryCode == myClientSite.FactoryCode
                                    select tb).Count();
                if (FactoryCodee==0)
                {
                    //写入数据
                    myClientSite.ClientID = ClientID; //保存进去的客户ID等于页面选中的客户ID

                    // 写入数据库
                    myModels.SYS_ClientSite.Add(myClientSite);
                    //保存数据库
                    myModels.SaveChanges();
                    strMsg = "success";
                }
                else
                {
                    strMsg = "failed";
                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存修改工厂地址
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateClientSite(SYS_ClientSite myClientSite)
        {
            string strMsg = "failed";
            try
            {
                SYS_ClientSite ClientSite = (from tbClientSite in myModels.SYS_ClientSite
                                             where tbClientSite.ClientSiteID == myClientSite.ClientSiteID
                                             select tbClientSite).Single();

                //写入数据
                ClientSite.ClientSiteID = myClientSite.ClientSiteID;
                ClientSite.ClientID = myClientSite.ClientID;
                ClientSite.ContactsName = myClientSite.ContactsName;
                ClientSite.ContactsPhone = myClientSite.ContactsPhone;
                ClientSite.ClientSite = myClientSite.ClientSite;
                ClientSite.GatedotID = myClientSite.GatedotID;
                ClientSite.FactoryName = myClientSite.FactoryName;
                ClientSite.FactoryCode = myClientSite.FactoryCode;
                ClientSite.Remark = myClientSite.Remark;

                myModels.Entry(ClientSite).State = System.Data.Entity.EntityState.Modified;
                myModels.SaveChanges();
                strMsg = "success";
                return Json(strMsg, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存删除工厂地址
        /// </summary>
        /// <returns></returns>
        public ActionResult DeleteClientSite(int ClientSiteID)
        {
            string strMsg = "failed";
            try
            {
                var listClientSite = (from tbClientSite in myModels.SYS_ClientSite
                                      where tbClientSite.ClientSiteID == ClientSiteID
                                      select tbClientSite).ToList();

                myModels.SYS_ClientSite.RemoveRange(listClientSite);
                myModels.SaveChanges();
                strMsg = "success";
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// 保存新增客户标准运费报价明细
        /// </summary>
        /// <returns></returns>
        public ActionResult SaveClientMeasure(SYS_OfferDetail myOfferDetail)
        {
            string strMsg = "failed";
            try
            {

                var OfferDetaill = (from tb in myModels.SYS_OfferDetail
                                    where tb.OfferID == myOfferDetail.OfferID && tb.HaulWayID == myOfferDetail.HaulWayID && tb.CabinetType.Trim() == myOfferDetail.CabinetType.Trim()
                                    select tb).Count();
                if (OfferDetaill==0)
                {
                    //写入数据
                    myOfferDetail.StaffID = Convert.ToInt32(Session["StaffID"].ToString()); //强势转换
                    myOfferDetail.Currency = "RMB";
                    myOfferDetail.ExpenseID = myModels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "运费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();
                    
                    myModels.SYS_OfferDetail.Add(myOfferDetail);
                    //保存数据库
                    myModels.SaveChanges();
                    strMsg = "success";
                }       
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
        
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// 保存修改客户标准运费报价明细
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateClientMeasure(SYS_OfferDetail myOfferDetail)
        {
            string strMsg = "failed";
            try
            {
                var OfferDetaill = (from tb in myModels.SYS_OfferDetail
                                    where tb.OfferID==myOfferDetail.OfferID && tb.HaulWayID == myOfferDetail.HaulWayID && tb.CabinetType.Trim() == myOfferDetail.CabinetType.Trim()
                                    select tb).ToList();
                if (OfferDetaill.Where(m=>m.OfferDetailID!=myOfferDetail.OfferDetailID).Count()==0)
                {
                    SYS_OfferDetail MOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                                    where tbOfferDetail.OfferDetailID == myOfferDetail.OfferDetailID
                                                    select tbOfferDetail).Single();

                    //写入数据
                    MOfferDetail.OfferDetailID = myOfferDetail.OfferDetailID;
                    MOfferDetail.HaulWayID = myOfferDetail.HaulWayID;
                    MOfferDetail.EtryClasses = myOfferDetail.EtryClasses;
                    MOfferDetail.CabinetType = myOfferDetail.CabinetType;
                    MOfferDetail.Money = myOfferDetail.Money;
                    MOfferDetail.Remark = myOfferDetail.Remark;
                    MOfferDetail.BoxQuantity = myOfferDetail.BoxQuantity;
                    MOfferDetail.StaffID = Convert.ToInt32(Session["StaffID"].ToString()); //强势转换

                    myModels.Entry(MOfferDetail).State = System.Data.Entity.EntityState.Modified;
                    myModels.SaveChanges();
                    strMsg = "success";
                    return Json(strMsg, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存删除选中的客户标准运费报价明细
        /// </summary>
        /// <returns></returns>
        public ActionResult DeleteClientMeasure(int OfferDetailID)
        {
            string strMsg = "failed";
            try
            {
                var listOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                      where tbOfferDetail.OfferDetailID == OfferDetailID
                                       select tbOfferDetail).ToList();

                myModels.SYS_OfferDetail.RemoveRange(listOfferDetail);
                myModels.SaveChanges();
                strMsg = "success";
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存客户应收运费上方表格的新增报价表
        /// </summary>
        /// <returns></returns>
        public ActionResult SaveReceivableOffer(SYS_Offer myOffer)
        {
            string strMsg = "failed";
            try
            {
                //实例化
                SYS_Offer Offer = new SYS_Offer();

                var ClientReceivable = myModels.SYS_Offer.Where(m => m.ClientID == myOffer.ClientID&& m.WhetherShui == myOffer.WhetherShui&&m.OfferType.Trim()== "客户应收费用").Count();
              
                if (ClientReceivable == 0 )
                {
                    //写入数据
                    myOffer.OfferType = "客户应收费用";

                    // 写入数据库
                    myModels.SYS_Offer.Add(myOffer);
                    //保存数据库
                    myModels.SaveChanges();
                    strMsg = "success";
                }
                else
                {
                    strMsg = "failed";
                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存客户应收运费上方表格的编辑报价表
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateReceivableOffer(SYS_Offer myOffer)
        {
            string strMsg = "failed";
            try
            {
                var ClientReceivable = myModels.SYS_Offer.Where(m => m.ClientID == myOffer.ClientID && m.WhetherShui == myOffer.WhetherShui&&m.OfferType.Trim()== "客户应收费用").Count();

                if (ClientReceivable == 0)
                {
                    var MOffer = (from tbOffer in myModels.SYS_Offer
                                        where tbOffer.OfferID == myOffer.OfferID
                                        select tbOffer).Single();

                    //写入数据
                    MOffer.ClientID = myOffer.ClientID;
                    MOffer.OfferDate = myOffer.OfferDate;
                    MOffer.TakeEffectDate = myOffer.TakeEffectDate;
                    MOffer.LoseEfficacyDate = myOffer.LoseEfficacyDate;
                    MOffer.WhetherShui = myOffer.WhetherShui;
                    MOffer.Remark = myOffer.Remark;

                    myModels.Entry(MOffer).State = System.Data.Entity.EntityState.Modified;
                    myModels.SaveChanges();
                    strMsg = "success";
                    return Json(strMsg, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    strMsg = "failed";
                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存删除客户应收运费上方表格的方法
        /// </summary>
        /// <returns></returns>
        public ActionResult DeleteReceivableOffer(int OfferID)
        {
            string strMsg = "failed";
            try
            {
                var listOffer = (from tbOffer in myModels.SYS_Offer
                                 where tbOffer.OfferID == OfferID
                                 select tbOffer).ToList();

                myModels.SYS_Offer.RemoveRange(listOffer);
                myModels.SaveChanges();
                strMsg = "success";
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// 保存客户应收运费报价明细表格数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SaveReceivableOfferDetail(SYS_OfferDetail myOfferDetail)
        {
            string strMsg = "failed";
            try
            {
                var OfferDetaill = (from tb in myModels.SYS_OfferDetail
                                    where tb.OfferID == myOfferDetail.OfferID && tb.HaulWayID == myOfferDetail.HaulWayID && tb.CabinetType.Trim() == myOfferDetail.CabinetType.Trim()
                                    select tb).Count();
                if (OfferDetaill == 0)
                {
                    //写入数据
                    myOfferDetail.StaffID = Convert.ToInt32(Session["StaffID"].ToString()); //强势转换
                    myOfferDetail.Currency = "RMB";
                    myOfferDetail.ExpenseID = myModels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "运费" && m.SettleAccounts.Trim() == "应收").Select(m => m.ExpenseID).Single();

                    myModels.SYS_OfferDetail.Add(myOfferDetail);
                    //保存数据库
                    myModels.SaveChanges();
                    strMsg = "success";
                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }

            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// 保存修改客户应收运费报价明细表格的数据
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateReceivableOfferDetail(SYS_OfferDetail myOfferDetail)
        {
            string strMsg = "failed";
            try
            {
                var OfferDetaill = (from tb in myModels.SYS_OfferDetail
                                    where tb.OfferID == myOfferDetail.OfferID && tb.HaulWayID == myOfferDetail.HaulWayID && tb.CabinetType.Trim() == myOfferDetail.CabinetType.Trim()
                                    select tb).ToList();
                if (OfferDetaill.Where(m => m.OfferDetailID != myOfferDetail.OfferDetailID).Count() == 0)
                {
                    SYS_OfferDetail MOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                                    where tbOfferDetail.OfferDetailID == myOfferDetail.OfferDetailID
                                                    select tbOfferDetail).Single();

                    //写入数据
                    MOfferDetail.OfferDetailID = myOfferDetail.OfferDetailID;
                    MOfferDetail.HaulWayID = myOfferDetail.HaulWayID;
                    MOfferDetail.EtryClasses = myOfferDetail.EtryClasses;
                    MOfferDetail.CabinetType = myOfferDetail.CabinetType;
                    MOfferDetail.Money = myOfferDetail.Money;
                    MOfferDetail.Remark = myOfferDetail.Remark;
                    MOfferDetail.BoxQuantity = myOfferDetail.BoxQuantity;
                    MOfferDetail.StaffID = Convert.ToInt32(Session["StaffID"].ToString()); //强势转换

                    myModels.Entry(MOfferDetail).State = System.Data.Entity.EntityState.Modified;
                    myModels.SaveChanges();
                    strMsg = "success";
                    return Json(strMsg, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存删除客户应收运费报价明细表格的数据
        /// </summary>
        /// <returns></returns>
        public ActionResult DeleteReceivableOfferDetail(int OfferDetailID)
        {
            string strMsg = "failed";
            try
            {
                var listOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                       where tbOfferDetail.OfferDetailID == OfferDetailID
                                       select tbOfferDetail).ToList();

                myModels.SYS_OfferDetail.RemoveRange(listOfferDetail);
                myModels.SaveChanges();
                strMsg = "success";
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存新增车队标准运费的上方表格报价数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SaveMotorcadeOffce(SYS_Offer myOffer)
        {
            string strMsg = "failed";
            try
            {
                //实例化
                SYS_Offer Offer = new SYS_Offer();

                var MotorcadeOffce = myModels.SYS_Offer.Where(m => m.ClientID == myOffer.ClientID && m.WhetherShui == myOffer.WhetherShui&&m.OfferType.Trim()== "车队标准运费").Count();

                if (MotorcadeOffce == 0)
                {
                    //写入数据
                    myOffer.OfferType = "车队标准运费";

                    // 写入数据库
                    myModels.SYS_Offer.Add(myOffer);
                    //保存数据库
                    myModels.SaveChanges();
                    strMsg = "success";
                }
                else
                {
                    strMsg = "failed";
                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存修改车队标准运费的上方表格报价数据
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateMotorcadeOffce(SYS_Offer myOffer)
        {
            string strMsg = "failed";
            try
            {
                var ClientReceivable = myModels.SYS_Offer.Where(m => m.ClientID == myOffer.ClientID && m.WhetherShui == myOffer.WhetherShui && m.OfferType.Trim() == "车队标准运费").Count();

                if (ClientReceivable == 0)
                {
                    var MOffer = (from tbOffer in myModels.SYS_Offer
                                  where tbOffer.OfferID == myOffer.OfferID
                                  select tbOffer).Single();

                    //写入数据
                    MOffer.ClientID = myOffer.ClientID;
                    MOffer.OfferDate = myOffer.OfferDate;
                    MOffer.TakeEffectDate = myOffer.TakeEffectDate;
                    MOffer.LoseEfficacyDate = myOffer.LoseEfficacyDate;
                    MOffer.WhetherShui = myOffer.WhetherShui;
                    MOffer.Remark = myOffer.Remark;

                    myModels.Entry(MOffer).State = System.Data.Entity.EntityState.Modified;
                    myModels.SaveChanges();
                    strMsg = "success";
                    return Json(strMsg, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    strMsg = "failed";
                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存删除车队标准运费的上方表格数据的报价
        /// </summary>
        /// <returns></returns>
        public ActionResult DeleteMotorcadeOffce(int OfferID)
        {
            string strMsg = "failed";
            try
            {
                var listOffer = (from tbOffer in myModels.SYS_Offer
                                 where tbOffer.OfferID == OfferID
                                 select tbOffer).ToList();

                myModels.SYS_Offer.RemoveRange(listOffer);
                myModels.SaveChanges();
                strMsg = "success";
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// 保存下方表格车队标准运费报价明细表
        /// </summary>
        /// <returns></returns>
        public ActionResult SaveMotorcadeMeasureOfferDetail(SYS_OfferDetail myOfferDetail)
        {
            string strMsg = "failed";
            try
            {
                var OfferDetaill = (from tb in myModels.SYS_OfferDetail
                                    where tb.OfferID == myOfferDetail.OfferID && tb.HaulWayID == myOfferDetail.HaulWayID && tb.CabinetType.Trim() == myOfferDetail.CabinetType.Trim()
                                    select tb).Count();
                if (OfferDetaill == 0)
                {
                    //写入数据
                    myOfferDetail.StaffID = Convert.ToInt32(Session["StaffID"].ToString()); //强势转换
                    myOfferDetail.Currency = "RMB";
                    myOfferDetail.ExpenseID = myModels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "运费" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();

                    myModels.SYS_OfferDetail.Add(myOfferDetail);
                    //保存数据库
                    myModels.SaveChanges();
                    strMsg = "success";
                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }

            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// 保存修改下方表格车队标准运费报价明细表
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateMotorcadeMeasureOfferDetail(SYS_OfferDetail myOfferDetail)
        {
            string strMsg = "failed";
            try
            {
                var OfferDetaill = (from tb in myModels.SYS_OfferDetail
                                    where tb.OfferID == myOfferDetail.OfferID && tb.HaulWayID == myOfferDetail.HaulWayID && tb.CabinetType.Trim() == myOfferDetail.CabinetType.Trim()
                                    select tb).ToList();
                if (OfferDetaill.Where(m => m.OfferDetailID != myOfferDetail.OfferDetailID).Count() == 0)
                {
                    SYS_OfferDetail MOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                                    where tbOfferDetail.OfferDetailID == myOfferDetail.OfferDetailID
                                                    select tbOfferDetail).Single();

                    //写入数据
                    MOfferDetail.OfferDetailID = myOfferDetail.OfferDetailID;
                    MOfferDetail.HaulWayID = myOfferDetail.HaulWayID;
                    MOfferDetail.EtryClasses = myOfferDetail.EtryClasses;
                    MOfferDetail.CabinetType = myOfferDetail.CabinetType;
                    MOfferDetail.Money = myOfferDetail.Money;
                    MOfferDetail.Remark = myOfferDetail.Remark;
                    MOfferDetail.BoxQuantity = myOfferDetail.BoxQuantity;
                    MOfferDetail.StaffID = Convert.ToInt32(Session["StaffID"].ToString()); //强势转换

                    myModels.Entry(MOfferDetail).State = System.Data.Entity.EntityState.Modified;
                    myModels.SaveChanges();
                    strMsg = "success";
                    return Json(strMsg, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存删除下方表格车队标准运费报价明细表
        /// </summary>
        /// <returns></returns>
        public ActionResult DeleteMotorcadeMeasureOfferDetail(int OfferDetailID)
        {
            string strMsg = "failed";
            try
            {
                var listOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                       where tbOfferDetail.OfferDetailID == OfferDetailID
                                       select tbOfferDetail).ToList();

                myModels.SYS_OfferDetail.RemoveRange(listOfferDetail);
                myModels.SaveChanges();
                strMsg = "success";
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存新增司机产值下方表格的报价明细表
        /// </summary>
        /// <param name="myOfferDetail"></param>
        /// <returns></returns>
        public ActionResult SaveChauffeurProduceOfferDetail(SYS_OfferDetail myOfferDetail)
        {
            string strMsg = "failed";
            try
            {

                var OfferDetaill = (from tb in myModels.SYS_OfferDetail
                                    where tb.OfferID == myOfferDetail.OfferID && tb.HaulWayID == myOfferDetail.HaulWayID && tb.CabinetType.Trim() == myOfferDetail.CabinetType.Trim()
                                    select tb).Count();
                if (OfferDetaill == 0)
                {
                    //写入数据
                    myOfferDetail.StaffID = Convert.ToInt32(Session["StaffID"].ToString()); //强势转换
                    myOfferDetail.Currency = "RMB";
                    myOfferDetail.ExpenseID = myModels.SYS_Expense.Where(m => m.ExpenseName.Trim() == "运费" && m.SettleAccounts.Trim() == "应付").Select(m => m.ExpenseID).Single();

                    myModels.SYS_OfferDetail.Add(myOfferDetail);
                    //保存数据库
                    myModels.SaveChanges();
                    strMsg = "success";
                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }

            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 保存修改司机产值下方表格的报价明细表
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateChauffeurProduceOfferDetail(SYS_OfferDetail myOfferDetail)
        {
            string strMsg = "failed";
            try
            {
                var OfferDetaill = (from tb in myModels.SYS_OfferDetail
                                    where tb.OfferID == myOfferDetail.OfferID && tb.HaulWayID == myOfferDetail.HaulWayID && tb.CabinetType.Trim() == myOfferDetail.CabinetType.Trim()
                                    select tb).ToList();
                if (OfferDetaill.Where(m => m.OfferDetailID != myOfferDetail.OfferDetailID).Count() == 0)
                {
                    SYS_OfferDetail MOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                                    where tbOfferDetail.OfferDetailID == myOfferDetail.OfferDetailID
                                                    select tbOfferDetail).Single();

                    //写入数据
                    MOfferDetail.OfferDetailID = myOfferDetail.OfferDetailID;
                    MOfferDetail.HaulWayID = myOfferDetail.HaulWayID;
                    MOfferDetail.EtryClasses = myOfferDetail.EtryClasses;
                    MOfferDetail.CabinetType = myOfferDetail.CabinetType;
                    MOfferDetail.Money = myOfferDetail.Money;
                    MOfferDetail.Remark = myOfferDetail.Remark;
                    MOfferDetail.BoxQuantity = myOfferDetail.BoxQuantity;
                    MOfferDetail.StaffID = Convert.ToInt32(Session["StaffID"].ToString()); //强势转换

                    myModels.Entry(MOfferDetail).State = System.Data.Entity.EntityState.Modified;
                    myModels.SaveChanges();
                    strMsg = "success";
                    return Json(strMsg, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// 保存删除选中的司机产值报价明细
        /// </summary>
        /// <returns></returns>
        public ActionResult DeleteChauffeurProduceOfferDetail(int OfferDetailID)
        {
            string strMsg = "failed";
            try
            {
                var listOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                       where tbOfferDetail.OfferDetailID == OfferDetailID
                                       select tbOfferDetail).ToList();

                myModels.SYS_OfferDetail.RemoveRange(listOfferDetail);
                myModels.SaveChanges();
                strMsg = "success";
            }
            catch (Exception)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }

        #endregion

        #region 打印

        /// <summary>
        /// 打印客户管理信息
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintAchievement(int? ClientID, string AAA, string ClientCode, string ClientAbbreviation, int? GatedotID, int? MessageID, bool? WhetherStart)
        {
            #region 数据查询
            var listClient = (from tbClient in myModels.SYS_Client
                              join tbGatedot in myModels.SYS_Gatedot on tbClient.GatedotID equals tbGatedot.GatedotID
                              join tbstaff in myModels.SYS_Staff on tbClient.StaffID equals tbstaff.StaffID
                              orderby tbClient.ClientID descending
                              select new ClientVo
                              {
                                  ClientID = tbClient.ClientID,//客户信息ID
                                  ClientCode = tbClient.ClientCode,//客户代码
                                  ClientAbbreviation = tbClient.ClientAbbreviation,//客户简称
                                  ChineseName = tbClient.ChineseName,//中文名称
                                  ClientRank = tbClient.ClientRank,//客户级别
                                  CustomsCode = tbClient.CustomsCode,//海关编码
                                  GatedotName = tbGatedot.GatedotName,//所属区域
                                  GatedotID = tbGatedot.GatedotID,//所属区域ID（门点ID）
                                  ClientSource = tbClient.ClientSource,//客户来源
                                  ClientPhone = tbClient.ClientPhone,//客户电话
                                  WhetherStart = tbClient.WhetherStart,//是否启用
                                  WhetherStartt = "",
                                  MessageID = tbClient.MessageID,//组织结构ID
                                  StaffName = tbstaff.StaffName,// 业务员
                                  ClientFax = tbClient.ClientFax,//客户传真
                                  Email = tbClient.Email,
                                  PostCode = tbClient.PostCode,//邮编
                                  OfficeHours1 = tbClient.OfficeHours.ToString(),//上班时间
                                  ClosingTime1 = tbClient.ClosingTime.ToString(),//下班时间
                                  Site = tbClient.Site,//地址
                                  Website = tbClient.Website,//网站
                                  OpenAccount = tbClient.OpenAccount,//开户行
                                  OpenAccountCode = tbClient.OpenAccountCode,//开户行账户
                                  Describe = tbClient.Describe,//描述
                              }).ToList();

            for (int i = 0; i < listClient.Count(); i++)  //转换true false 为是否
            {
                if (listClient[i].WhetherStart == true)
                {
                    listClient[i].WhetherStartt = "是";
                }
                else
                {
                    listClient[i].WhetherStartt = "否";
                }
            }

            if (AAA == "AAA")//页面传过来的参数
            {
                listClient = listClient.Where(m => m.ClientID == ClientID).ToList();
                return Json(listClient, JsonRequestBehavior.AllowGet);
            }
            if (!string.IsNullOrEmpty(ClientCode))//客户代码
            {
                ClientCode = ClientCode.Trim();
                listClient = listClient.Where(p => p.ClientCode.Contains(ClientCode)).ToList();
            }
            if (!string.IsNullOrEmpty(ClientAbbreviation))//客户简称
            {
                ClientAbbreviation = ClientAbbreviation.Trim();
                listClient = listClient.Where(p => p.ClientAbbreviation.Contains(ClientAbbreviation)).ToList();
            }
            if (GatedotID > 0)//所属区域ID（门点ID）
            {
                listClient = listClient.Where(m => m.GatedotID == GatedotID).ToList();
            }
            if (MessageID > 0)//组织结构ID
            {
                listClient = listClient.Where(m => m.MessageID == MessageID).ToList();
            }
            if (WhetherStart != null)//是否启用  只要判断不为空就可以了
            {
                listClient = listClient.Where(m => m.WhetherStart == WhetherStart).ToList();
            }
            #endregion

            //将查询出来的结果转化为DataTable的格式
            DataTable dtResult = LINQToDataTable(listClient);

            //1、实例化数据集
            PrintReport.ClientDB dbReport = new PrintReport.ClientDB();
            //2、将dtResult放入数据集中名为"tbAchievement"的表格中
            dbReport.Tables["tabClient"].Merge(dtResult);
            //3、实例化数据报表
            PrintReport.Client rp = new PrintReport.Client();
            //4、获取报表的物理文件路径
            string strRpPath = System.Web.HttpContext.Current.Server.MapPath("~/") + "Areas\\ClientManage\\PrintReport\\Client.rpt";
            //5、将报表加载到报表模板中
            rp.Load(strRpPath);
            //6、设置报表的数据源
            rp.SetDataSource(dbReport);
            //7、把报表转化为文件流输出
            Stream dbStream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);

            //返回
            return File(dbStream, "application/pdf");

        }

        /// <summary>
        /// 打印客户标准运费报价明细
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintClientMeasure(int? OfferID,string myArray)
        {
            #region 查询数据
            var listOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                   join tbExpense in myModels.SYS_Expense on tbOfferDetail.ExpenseID equals tbExpense.ExpenseID
                                   join tbHaulWay in myModels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWay.HaulWayID
                                   join tbMention in myModels.SYS_Mention on tbHaulWay.MentionAreaID equals tbMention.MentionID
                                   join tbLoading in myModels.SYS_Gatedot on tbHaulWay.LoadingAreaID equals tbLoading.GatedotID
                                   join tbAlsoTank in myModels.SYS_Mention on tbHaulWay.AlsoTankAreaID equals tbAlsoTank.MentionID
                                   join tbStaff in myModels.SYS_Staff on tbOfferDetail.StaffID equals tbStaff.StaffID
                                   orderby tbOfferDetail.OfferDetailID descending
                                   select new ClientVo
                                   {
                                       HaulWayID = tbHaulWay.HaulWayID,
                                       OfferID = tbOfferDetail.OfferID,
                                       OfferDetailID = tbOfferDetail.OfferDetailID,
                                       Abbreviation = tbMention.Abbreviation.Trim(),
                                       Abbreviationn = tbAlsoTank.Abbreviation.Trim(),
                                       GatedotName = tbLoading.GatedotName.Trim(),
                                       HaulWayDescription = tbMention.Abbreviation.Trim() + "-" + tbLoading.GatedotName.Trim() + "-" + tbAlsoTank.Abbreviation.Trim(),
                                       ExpenseName = tbExpense.ExpenseName,
                                       CabinetType = tbOfferDetail.CabinetType.Trim(),
                                       EtryClasses = tbOfferDetail.EtryClasses.Trim(),
                                       Money = tbOfferDetail.Money,
                                       Currency = tbOfferDetail.Currency,
                                       Remark = tbOfferDetail.Remark,
                                       StaffName = tbStaff.StaffName,
                                       BoxQuantity = tbOfferDetail.BoxQuantity
                                   }).ToList();
            if (OfferID > 0)
            {
                listOfferDetail = listOfferDetail.Where(m => m.OfferID == OfferID).ToList();
            }

            if (!string.IsNullOrEmpty(myArray))
            {
                string[] AAA = myArray.Split(',');//用英文状态下的逗号分割字符串
                for (int i = 0; i < listOfferDetail.Count(); i++)
                {
                    var F = listOfferDetail[i].OfferDetailID.ToString();
                    int id = Array.IndexOf(AAA, F);
                    if (id == -1)
                    {
                        listOfferDetail.Remove(listOfferDetail[i]);
                        i = 0;
                    }
                }
            }
            #endregion

            //将查询出来的结果转化为DataTable的格式
            DataTable dtResult = LINQToDataTable(listOfferDetail);

            //1、实例化数据集
            PrintReport.ClientDB dbReport = new PrintReport.ClientDB();
            //2、将dtResult放入数据集中名为"tbAchievement"的表格中
            dbReport.Tables["ClientMeasure"].Merge(dtResult);
            //3、实例化数据报表
            PrintReport.ClientMeasure rp = new PrintReport.ClientMeasure();
            //4、获取报表的物理文件路径
            string strRpPath = System.Web.HttpContext.Current.Server.MapPath("~/") + "Areas\\ClientManage\\PrintReport\\ClientMeasure.rpt";
            //5、将报表加载到报表模板中
            rp.Load(strRpPath);
            //6、设置报表的数据源
            rp.SetDataSource(dbReport);
            //7、把报表转化为文件流输出
            Stream dbStream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);

            //返回
            return File(dbStream, "application/pdf");
        }


        /// <summary>
        /// 打印客户应收运费的报价表格
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintClientReceivable(int? OfferID,string myArray)
        {
            #region 数据查询

            var listOffer = (from tbOffer in myModels.SYS_Offer
                             join tbClient in myModels.SYS_Client on tbOffer.ClientID equals tbClient.ClientID
                             where tbOffer.OfferType.Trim() == "客户应收费用"
                             orderby tbOffer.OfferID descending
                             select new ClientVo
                             {
                                 OfferID = tbOffer.OfferID,
                                 OfferDate1 = tbOffer.OfferDate.ToString(),
                                 OfferDate = tbOffer.OfferDate,
                                 TakeEffectDate1 = tbOffer.TakeEffectDate.ToString(),
                                 LoseEfficacyDate1 = tbOffer.LoseEfficacyDate.ToString(),
                                 WhetherShuii = "",
                                 WhetherShui = tbOffer.WhetherShui,
                                 Remark = tbOffer.Remark,
                                 ClientAbbreviation = tbClient.ClientAbbreviation,
                                 ClientCode = tbClient.ClientCode,
                                 ClientID = tbClient.ClientID,

                             }).ToList();
            for (int i = 0; i < listOffer.Count(); i++)  //转换true false 为是否
            {
                if (listOffer[i].WhetherShui == true)
                {
                    listOffer[i].WhetherShuii = "是";
                }
                else
                {
                    listOffer[i].WhetherShuii = "否";
                }
            }

            if (!string.IsNullOrEmpty(myArray))
            {
                string[] AAA = myArray.Split(',');//用英文状态下的逗号分割字符串
                for (int i = 0; i < listOffer.Count(); i++)
                {
                    var F = listOffer[i].OfferID.ToString();
                    int id = Array.IndexOf(AAA, F);
                    if (id == -1)
                    {
                        listOffer.Remove(listOffer[i]);
                        i = 0;
                    }
                }
            }

            #endregion

            //将查询出来的结果转化为DataTable的格式
            DataTable dtResult = LINQToDataTable(listOffer);

            //1、实例化数据集
            PrintReport.ClientDB dbReport = new PrintReport.ClientDB();
            //2、将dtResult放入数据集中名为"tbAchievement"的表格中
            dbReport.Tables["tabClientReceivable"].Merge(dtResult);
            //3、实例化数据报表
            PrintReport.ClientReceivable rp = new PrintReport.ClientReceivable();
            //4、获取报表的物理文件路径
            string strRpPath = System.Web.HttpContext.Current.Server.MapPath("~/") + "Areas\\ClientManage\\PrintReport\\ClientReceivable.rpt";
            //5、将报表加载到报表模板中
            rp.Load(strRpPath);
            //6、设置报表的数据源
            rp.SetDataSource(dbReport);
            //7、把报表转化为文件流输出
            Stream dbStream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);

            //返回
            return File(dbStream, "application/pdf");
        }

        /// <summary>
        /// 打印客户应收运费的报价明细表
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintReceivableOfferDetail(int? OfferID, string myArray)
        {
            #region 数据查询
            var listOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                   join tbExpense in myModels.SYS_Expense on tbOfferDetail.ExpenseID equals tbExpense.ExpenseID
                                   join tbHaulWay in myModels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWay.HaulWayID
                                   join tbMention in myModels.SYS_Mention on tbHaulWay.MentionAreaID equals tbMention.MentionID
                                   join tbLoading in myModels.SYS_Gatedot on tbHaulWay.LoadingAreaID equals tbLoading.GatedotID
                                   join tbAlsoTank in myModels.SYS_Mention on tbHaulWay.AlsoTankAreaID equals tbAlsoTank.MentionID
                                   join tbStaff in myModels.SYS_Staff on tbOfferDetail.StaffID equals tbStaff.StaffID
                                   orderby tbOfferDetail.OfferDetailID descending
                                   select new ClientVo
                                   {
                                       HaulWayID = tbHaulWay.HaulWayID,
                                       OfferID = tbOfferDetail.OfferID,
                                       OfferDetailID = tbOfferDetail.OfferDetailID,
                                       Abbreviation = tbMention.Abbreviation.Trim(),
                                       Abbreviationn = tbAlsoTank.Abbreviation.Trim(),
                                       GatedotName = tbLoading.GatedotName.Trim(),
                                       HaulWayDescription = tbMention.Abbreviation.Trim() + "-" + tbLoading.GatedotName.Trim() + "-" + tbAlsoTank.Abbreviation.Trim(),
                                       ExpenseName = tbExpense.ExpenseName,
                                       CabinetType = tbOfferDetail.CabinetType.Trim(),
                                       EtryClasses = tbOfferDetail.EtryClasses.Trim(),
                                       Money = tbOfferDetail.Money,
                                       Currency = tbOfferDetail.Currency,
                                       Remark = tbOfferDetail.Remark,
                                       StaffName = tbStaff.StaffName,
                                       BoxQuantity = tbOfferDetail.BoxQuantity,
                                   }).ToList();

            if (OfferID > 0)
            {
                listOfferDetail = listOfferDetail.Where(m => m.OfferID == OfferID).ToList();
            }

            if (!string.IsNullOrEmpty(myArray))
            {
                string[] AAA = myArray.Split(',');//用英文状态下的逗号分割字符串
                for (int i = 0; i < listOfferDetail.Count(); i++)
                {
                    var F = listOfferDetail[i].OfferDetailID.ToString();
                    int id = Array.IndexOf(AAA, F);
                    if (id == -1)
                    {
                        listOfferDetail.Remove(listOfferDetail[i]);
                        i = 0;
                    }
                }
            }
            #endregion

            //将查询出来的结果转化为DataTable的格式
            DataTable dtResult = LINQToDataTable(listOfferDetail);

            //1、实例化数据集
            PrintReport.ClientDB dbReport = new PrintReport.ClientDB();
            //2、将dtResult放入数据集中名为"tbAchievement"的表格中
            dbReport.Tables["tabOfferDetail"].Merge(dtResult);
            //3、实例化数据报表
            PrintReport.ClientOfferDetail rp = new PrintReport.ClientOfferDetail();
            //4、获取报表的物理文件路径
            string strRpPath = System.Web.HttpContext.Current.Server.MapPath("~/") + "Areas\\ClientManage\\PrintReport\\ClientOfferDetail.rpt";
            //5、将报表加载到报表模板中
            rp.Load(strRpPath);
            //6、设置报表的数据源
            rp.SetDataSource(dbReport);
            //7、把报表转化为文件流输出
            Stream dbStream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);

            //返回
            return File(dbStream, "application/pdf");
        }

        /// <summary>
        /// 打印车队标准运费的报价表
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintMotorcadeOffce(int? ClientID, int? OfferID, string ClientAbbreviation, string OfferDate, string OfferDatee, bool? WhetherShui, string myArray)
        {
            #region 数据查询
            var listOffer = (from tbOffer in myModels.SYS_Offer
                             join tbClient in myModels.SYS_Client on tbOffer.ClientID equals tbClient.ClientID
                             where tbOffer.OfferType.Trim() == "车队标准运费"
                             orderby tbOffer.OfferID descending
                             select new ClientVo
                             {
                                 OfferID = tbOffer.OfferID,
                                 OfferDate1 = tbOffer.OfferDate.ToString(),
                                 OfferDate = tbOffer.OfferDate,
                                 TakeEffectDate1 = tbOffer.TakeEffectDate.ToString(),
                                 LoseEfficacyDate1 = tbOffer.LoseEfficacyDate.ToString(),
                                 WhetherShuii = "",
                                 WhetherShui = tbOffer.WhetherShui,
                                 Remark = tbOffer.Remark,
                                 ClientAbbreviation = tbClient.ClientAbbreviation,
                                 ClientCode = tbClient.ClientCode,
                                 ClientID = tbClient.ClientID,

                             }).ToList();

            for (int i = 0; i < listOffer.Count(); i++)  //转换true false 为是否
            {
                if (listOffer[i].WhetherShui == true)
                {
                    listOffer[i].WhetherShuii = "是";
                }
                else
                {
                    listOffer[i].WhetherShuii = "否";
                }
            }

            if (ClientID > 0)
            {
                listOffer = listOffer.Where(m => m.ClientID == ClientID).ToList();
            }
            if (!string.IsNullOrEmpty(ClientAbbreviation))//客户简称
            {
                ClientAbbreviation = ClientAbbreviation.Trim();
                listOffer = listOffer.Where(p => p.ClientAbbreviation.Contains(ClientAbbreviation)).ToList();
            }

            //判断时间段去查询的方法，有三种情况，所以需要写三个判断
            if (!string.IsNullOrEmpty(OfferDate) && string.IsNullOrEmpty(OfferDatee))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(OfferDate);
                    listOffer = listOffer.Where(m => m.OfferDate == Time).ToList();
                }
                catch (Exception)
                {
                    listOffer = listOffer.Where(m => m.OfferID == OfferID).ToList();
                }
            }
            else if (string.IsNullOrEmpty(OfferDate) && !string.IsNullOrEmpty(OfferDatee))
            {
                try
                {
                    DateTime Times = Convert.ToDateTime(OfferDatee);
                    listOffer = listOffer.Where(m => m.OfferDate <= Times).ToList();
                }
                catch (Exception)
                {
                    listOffer = listOffer.Where(m => m.OfferID == OfferID).ToList();
                }
            }
            else if (!string.IsNullOrEmpty(OfferDate) && !string.IsNullOrEmpty(OfferDatee))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(OfferDate);
                    DateTime Times = Convert.ToDateTime(OfferDatee);
                    listOffer = listOffer.Where(m => m.OfferDate >= Time && m.OfferDate <= Times).ToList();
                }
                catch (Exception)
                {
                    listOffer = listOffer.Where(m => m.OfferID == OfferID).ToList();
                }
            }

            if (WhetherShui != null)//是否含税  只要判断不为空就可以了
            {
                listOffer = listOffer.Where(m => m.WhetherShui == WhetherShui).ToList();
            }

            if (!string.IsNullOrEmpty(myArray))
            {
                string[] AAA = myArray.Split(',');//用英文状态下的逗号分割字符串
                for (int i = 0; i < listOffer.Count(); i++)
                {
                    var F = listOffer[i].OfferID.ToString();
                    int id = Array.IndexOf(AAA, F);
                    if (id == -1)
                    {
                        listOffer.Remove(listOffer[i]);
                        i = 0;
                    }
                }
            }

            #endregion

            //将查询出来的结果转化为DataTable的格式
            DataTable dtResult = LINQToDataTable(listOffer);

            //1、实例化数据集
            PrintReport.ClientDB dbReport = new PrintReport.ClientDB();
            //2、将dtResult放入数据集中名为"tbAchievement"的表格中
            dbReport.Tables["tabClientReceivable"].Merge(dtResult);
            //3、实例化数据报表
            PrintReport.MotorcadeOffce rp = new PrintReport.MotorcadeOffce();
            //4、获取报表的物理文件路径
            string strRpPath = System.Web.HttpContext.Current.Server.MapPath("~/") + "Areas\\ClientManage\\PrintReport\\MotorcadeOffce.rpt";
            //5、将报表加载到报表模板中
            rp.Load(strRpPath);
            //6、设置报表的数据源
            rp.SetDataSource(dbReport);
            //7、把报表转化为文件流输出
            Stream dbStream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);

            //返回
            return File(dbStream, "application/pdf");
        }

        /// <summary>
        /// 打印下方表格的车队标准运费报价明细表
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintMotorcadeMeasureOfferDetail(int? OfferID, int? OfferDetailID, string myArray)
        {
            #region 数据查询
            var listOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                   join tbExpense in myModels.SYS_Expense on tbOfferDetail.ExpenseID equals tbExpense.ExpenseID
                                   join tbHaulWay in myModels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWay.HaulWayID
                                   join tbMention in myModels.SYS_Mention on tbHaulWay.MentionAreaID equals tbMention.MentionID
                                   join tbLoading in myModels.SYS_Gatedot on tbHaulWay.LoadingAreaID equals tbLoading.GatedotID
                                   join tbAlsoTank in myModels.SYS_Mention on tbHaulWay.AlsoTankAreaID equals tbAlsoTank.MentionID
                                   join tbStaff in myModels.SYS_Staff on tbOfferDetail.StaffID equals tbStaff.StaffID
                                   orderby tbOfferDetail.OfferDetailID descending
                                   select new ClientVo
                                   {
                                       HaulWayID = tbHaulWay.HaulWayID,
                                       OfferID = tbOfferDetail.OfferID,
                                       OfferDetailID = tbOfferDetail.OfferDetailID,
                                       Abbreviation = tbMention.Abbreviation.Trim(),
                                       Abbreviationn = tbAlsoTank.Abbreviation.Trim(),
                                       GatedotName = tbLoading.GatedotName.Trim(),
                                       HaulWayDescription = tbMention.Abbreviation.Trim() + "-" + tbLoading.GatedotName.Trim() + "-" + tbAlsoTank.Abbreviation.Trim(),
                                       ExpenseName = tbExpense.ExpenseName,
                                       CabinetType = tbOfferDetail.CabinetType.Trim(),
                                       EtryClasses = tbOfferDetail.EtryClasses.Trim(),
                                       Money = tbOfferDetail.Money,
                                       Currency = tbOfferDetail.Currency,
                                       Remark = tbOfferDetail.Remark,
                                       StaffName = tbStaff.StaffName,
                                       BoxQuantity = tbOfferDetail.BoxQuantity,
                                   }).ToList();

            if (OfferID > 0)
            {
                listOfferDetail = listOfferDetail.Where(m => m.OfferID == OfferID).ToList();
            }
            if (!string.IsNullOrEmpty(myArray))
            {
                string[] AAA = myArray.Split(',');//用英文状态下的逗号分割字符串
                for (int i = 0; i < listOfferDetail.Count(); i++)
                {
                    var F = listOfferDetail[i].OfferDetailID.ToString();
                    int Id = Array.IndexOf(AAA, F);
                    if (Id==-1)
                    {
                        listOfferDetail.Remove(listOfferDetail[i]);
                        i = 0;
                    }
                }
            }

            #endregion

            //将查询出来的结果转化为DataTable的格式
            DataTable dtResult = LINQToDataTable(listOfferDetail);

            //1、实例化数据集
            PrintReport.ClientDB dbReport = new PrintReport.ClientDB();
            //2、将dtResult放入数据集中名为"tbAchievement"的表格中
            dbReport.Tables["tabOfferDetail"].Merge(dtResult);
            //3、实例化数据报表
            PrintReport.MotorcadeOfferDetail rp = new PrintReport.MotorcadeOfferDetail();
            //4、获取报表的物理文件路径
            string strRpPath = System.Web.HttpContext.Current.Server.MapPath("~/") + "Areas\\ClientManage\\PrintReport\\MotorcadeOfferDetail.rpt";
            //5、将报表加载到报表模板中
            rp.Load(strRpPath);
            //6、设置报表的数据源
            rp.SetDataSource(dbReport);
            //7、把报表转化为文件流输出
            Stream dbStream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);

            //返回
            return File(dbStream, "application/pdf");

        }


        /// <summary>
        /// 打印司机产值报价明细
        /// </summary>
        /// <returns></returns>
        public ActionResult PrintChauffeurProduce(int? OfferID, string myArray)
        {
            #region 查询数据
            var listOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                   join tbExpense in myModels.SYS_Expense on tbOfferDetail.ExpenseID equals tbExpense.ExpenseID
                                   join tbHaulWay in myModels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWay.HaulWayID
                                   join tbMention in myModels.SYS_Mention on tbHaulWay.MentionAreaID equals tbMention.MentionID
                                   join tbLoading in myModels.SYS_Gatedot on tbHaulWay.LoadingAreaID equals tbLoading.GatedotID
                                   join tbAlsoTank in myModels.SYS_Mention on tbHaulWay.AlsoTankAreaID equals tbAlsoTank.MentionID
                                   join tbStaff in myModels.SYS_Staff on tbOfferDetail.StaffID equals tbStaff.StaffID
                                   orderby tbOfferDetail.OfferDetailID descending
                                   select new ClientVo
                                   {
                                       HaulWayID = tbHaulWay.HaulWayID,
                                       OfferID = tbOfferDetail.OfferID,
                                       OfferDetailID = tbOfferDetail.OfferDetailID,
                                       Abbreviation = tbMention.Abbreviation,
                                       Abbreviationn = tbAlsoTank.Abbreviation,
                                       GatedotName = tbLoading.GatedotName,
                                       HaulWayDescription = tbMention.Abbreviation.Trim() + "-" + tbLoading.GatedotName.Trim() + "-" + tbAlsoTank.Abbreviation.Trim(),
                                       ExpenseName = tbExpense.ExpenseName,
                                       CabinetType = tbOfferDetail.CabinetType.Trim(),
                                       EtryClasses = tbOfferDetail.EtryClasses.Trim(),
                                       Money = tbOfferDetail.Money,
                                       Currency = tbOfferDetail.Currency,
                                       Remark = tbOfferDetail.Remark,
                                       StaffName = tbStaff.StaffName,
                                       BoxQuantity = tbOfferDetail.BoxQuantity
                                   }).ToList();

            if (OfferID > 0)
            {
                listOfferDetail = listOfferDetail.Where(m => m.OfferID == OfferID).ToList();
            }

            if (!string.IsNullOrEmpty(myArray))
            {
                string[] AAA = myArray.Split(',');//用英文状态下的逗号分割字符串
                for (int i = 0; i < listOfferDetail.Count(); i++)
                {
                    var F = listOfferDetail[i].OfferDetailID.ToString();
                    int id = Array.IndexOf(AAA, F);
                    if (id == -1)
                    {
                        listOfferDetail.Remove(listOfferDetail[i]);
                        i = 0;
                    }
                }
            }
            #endregion

            //将查询出来的结果转化为DataTable的格式
            DataTable dtResult = LINQToDataTable(listOfferDetail);

            //1、实例化数据集
            PrintReport.ClientDB dbReport = new PrintReport.ClientDB();
            //2、将dtResult放入数据集中名为"tbAchievement"的表格中
            dbReport.Tables["tabChauffeurProduce"].Merge(dtResult);
            //3、实例化数据报表
            PrintReport.ChauffeurProduce rp = new PrintReport.ChauffeurProduce();
            //4、获取报表的物理文件路径
            string strRpPath = System.Web.HttpContext.Current.Server.MapPath("~/") + "Areas\\ClientManage\\PrintReport\\ChauffeurProduce.rpt";
            //5、将报表加载到报表模板中
            rp.Load(strRpPath);
            //6、设置报表的数据源
            rp.SetDataSource(dbReport);
            //7、把报表转化为文件流输出
            Stream dbStream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);

            //返回
            return File(dbStream, "application/pdf");
        }

        public DataTable LINQToDataTable<T>(IEnumerable<T> varlist)
        {
            //定义要返回的DataTable的对象
            DataTable dtReturn = new DataTable();
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
                        //得到属性的类型
                        Type colType = pi.PropertyType;
                        //如果属性为泛型类型
                        if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition() == typeof(Nullable<>)))
                        {
                            //获取泛型类型的参数
                            colType = colType.GetGenericArguments()[0];
                        }
                        //将类型的属性名称与属性类型作为DataTable 的列数据
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

        #region 导出

        /// <summary>
        /// 导出客户信息
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportClientInfor(int? ClientID, string AAA, string ClientCode, string ClientAbbreviation, int? GatedotID, int? MessageID, bool? WhetherStart)
        {
            var listClient = (from tbClient in myModels.SYS_Client
                              join tbGatedot in myModels.SYS_Gatedot on tbClient.GatedotID equals tbGatedot.GatedotID
                              join tbstaff in myModels.SYS_Staff on tbClient.StaffID equals tbstaff.StaffID
                              orderby tbClient.ClientID descending
                              select new ClientVo
                              {
                                  ClientID = tbClient.ClientID,//客户信息ID
                                  ClientCode = tbClient.ClientCode,//客户代码
                                  ClientAbbreviation = tbClient.ClientAbbreviation,//客户简称
                                  ChineseName = tbClient.ChineseName,//中文名称
                                  ClientRank = tbClient.ClientRank,//客户级别
                                  CustomsCode = tbClient.CustomsCode,//海关编码
                                  GatedotName = tbGatedot.GatedotName,//所属区域
                                  GatedotID = tbGatedot.GatedotID,//所属区域ID（门点ID）
                                  ClientSource = tbClient.ClientSource,//客户来源
                                  ClientPhone = tbClient.ClientPhone,//客户电话
                                  WhetherStart = tbClient.WhetherStart,//是否启用
                                  WhetherStartt = "",
                                  MessageID = tbClient.MessageID,//组织结构ID
                                  StaffName = tbstaff.StaffName,// 业务员
                                  ClientFax = tbClient.ClientFax,//客户传真
                                  Email = tbClient.Email,
                                  PostCode = tbClient.PostCode,//邮编
                                  OfficeHours1 = tbClient.OfficeHours.ToString(),//上班时间
                                  ClosingTime1 = tbClient.ClosingTime.ToString(),//下班时间
                                  Site = tbClient.Site,//地址
                                  Website = tbClient.Website,//网站
                                  OpenAccount = tbClient.OpenAccount,//开户行
                                  OpenAccountCode = tbClient.OpenAccountCode,//开户行账户
                                  Describe = tbClient.Describe,//描述
                              }).ToList();

            for (int i = 0; i < listClient.Count(); i++)  //转换true false 为是否
            {
                if (listClient[i].WhetherStart == true)
                {
                    listClient[i].WhetherStartt = "是";
                }
                else
                {
                    listClient[i].WhetherStartt = "否";
                }
            }

            if (AAA == "AAA")//页面传过来的参数
            {
                listClient = listClient.Where(m => m.ClientID == ClientID).ToList();
                return Json(listClient, JsonRequestBehavior.AllowGet);
            }
            if (!string.IsNullOrEmpty(ClientCode))//客户代码
            {
                ClientCode = ClientCode.Trim();
                listClient = listClient.Where(p => p.ClientCode.Contains(ClientCode)).ToList();
            }
            if (!string.IsNullOrEmpty(ClientAbbreviation))//客户简称
            {
                ClientAbbreviation = ClientAbbreviation.Trim();
                listClient = listClient.Where(p => p.ClientAbbreviation.Contains(ClientAbbreviation)).ToList();
            }
            if (GatedotID > 0)//所属区域ID（门点ID）
            {
                listClient = listClient.Where(m => m.GatedotID == GatedotID).ToList();
            }
            if (MessageID > 0)//组织结构ID
            {
                listClient = listClient.Where(m => m.MessageID == MessageID).ToList();
            }
            if (WhetherStart != null)//是否启用  只要判断不为空就可以了
            {
                listClient = listClient.Where(m => m.WhetherStart == WhetherStart).ToList();
            }

            //Excel表格的创建步骤
            //第一步：创建Excel对象
            NPOI.HSSF.UserModel.HSSFWorkbook exBook = new NPOI.HSSF.UserModel.HSSFWorkbook();
            //第二步：创建Excel对象的工作簿
            NPOI.SS.UserModel.ISheet sheet = exBook.CreateSheet();
            //第三步：Excel表头设置
            //给sheet添加第一行的头部标题
            NPOI.SS.UserModel.IRow row1 = sheet.CreateRow(0);//创建行
            row1.CreateCell(0).SetCellValue("客户代码");
            row1.CreateCell(1).SetCellValue("客户简称");
            row1.CreateCell(2).SetCellValue("客户类型");
            row1.CreateCell(3).SetCellValue("中文名称");
            row1.CreateCell(4).SetCellValue("客户级别");
            row1.CreateCell(5).SetCellValue("海关编码");
            row1.CreateCell(6).SetCellValue("所属区域");
            row1.CreateCell(7).SetCellValue("客户来源");
            row1.CreateCell(8).SetCellValue("客户电话");
            row1.CreateCell(9).SetCellValue("业务员");
            row1.CreateCell(10).SetCellValue("上班时间");
            row1.CreateCell(11).SetCellValue("下班时间");
            row1.CreateCell(12).SetCellValue("是否启用");
            //4、循环写入数据
            for (var i = 0; i < listClient.Count(); i++)
            {
                NPOI.SS.UserModel.IRow rowTemp = sheet.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listClient[i].ClientCode);
                rowTemp.CreateCell(1).SetCellValue(listClient[i].ClientAbbreviation);
                rowTemp.CreateCell(2).SetCellValue(listClient[i].ClientType);
                rowTemp.CreateCell(3).SetCellValue(listClient[i].ChineseName);
                rowTemp.CreateCell(4).SetCellValue(listClient[i].ClientRank);
                rowTemp.CreateCell(5).SetCellValue(listClient[i].CustomsCode);
                rowTemp.CreateCell(6).SetCellValue(listClient[i].GatedotName);
                rowTemp.CreateCell(7).SetCellValue(listClient[i].ClientSource);
                rowTemp.CreateCell(8).SetCellValue(listClient[i].ClientPhone);
                rowTemp.CreateCell(9).SetCellValue(listClient[i].StaffName);
                rowTemp.CreateCell(10).SetCellValue(listClient[i].OfficeHours1);
                rowTemp.CreateCell(11).SetCellValue(listClient[i].ClosingTime1);
                rowTemp.CreateCell(12).SetCellValue(listClient[i].WhetherStartt);
            }
            //6、文件名
            var fileName = "广东信息科技有限公司客户管理信息报表" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            //5、将Excel表格转化为文件流输出
            MemoryStream exStream = new MemoryStream();
            exBook.Write(exStream);
            //输出之前调用Seek（偏移量，游标位置）方法：确定流开始的位置
            exStream.Seek(0, SeekOrigin.Begin);
            //6、文件名


            return File(exStream, "application/vnd.ms-excel", fileName);
        }

        /// <summary>
        /// 导出客户标准运费报价明细表
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportClientMeasure(int? OfferID, string myArray)
        {
            var listOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                   join tbExpense in myModels.SYS_Expense on tbOfferDetail.ExpenseID equals tbExpense.ExpenseID
                                   join tbHaulWay in myModels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWay.HaulWayID
                                   join tbMention in myModels.SYS_Mention on tbHaulWay.MentionAreaID equals tbMention.MentionID
                                   join tbLoading in myModels.SYS_Gatedot on tbHaulWay.LoadingAreaID equals tbLoading.GatedotID
                                   join tbAlsoTank in myModels.SYS_Mention on tbHaulWay.AlsoTankAreaID equals tbAlsoTank.MentionID
                                   join tbStaff in myModels.SYS_Staff on tbOfferDetail.StaffID equals tbStaff.StaffID
                                   orderby tbOfferDetail.OfferDetailID descending
                                   select new ClientVo
                                   {
                                       HaulWayID = tbHaulWay.HaulWayID,
                                       OfferID = tbOfferDetail.OfferID,
                                       OfferDetailID = tbOfferDetail.OfferDetailID,
                                       Abbreviation = tbMention.Abbreviation.Trim(),
                                       Abbreviationn = tbAlsoTank.Abbreviation.Trim(),
                                       GatedotName = tbLoading.GatedotName.Trim(),
                                       HaulWayDescription = tbMention.Abbreviation.Trim() + "-" + tbLoading.GatedotName.Trim() + "-" + tbAlsoTank.Abbreviation.Trim(),
                                       ExpenseName = tbExpense.ExpenseName,
                                       CabinetType = tbOfferDetail.CabinetType.Trim(),
                                       EtryClasses = tbOfferDetail.EtryClasses.Trim(),
                                       Money = tbOfferDetail.Money,
                                       Currency = tbOfferDetail.Currency,
                                       Remark = tbOfferDetail.Remark,
                                       StaffName = tbStaff.StaffName,
                                       BoxQuantity = tbOfferDetail.BoxQuantity
                                   }).ToList();
            if (OfferID > 0)
            {
                listOfferDetail = listOfferDetail.Where(m => m.OfferID == OfferID).ToList();
            }

            if (!string.IsNullOrEmpty(myArray))
            {
                string[] AAA = myArray.Split(',');//用英文状态下的逗号分割字符串
                for (int i = 0; i < listOfferDetail.Count(); i++)
                {
                    var F = listOfferDetail[i].OfferDetailID.ToString();
                    int id = Array.IndexOf(AAA, F);
                    if (id == -1)
                    {
                        listOfferDetail.Remove(listOfferDetail[i]);
                        i = 0;
                    }
                }
            }

            //Excel表格的创建步骤
            //第一步：创建Excel对象
            NPOI.HSSF.UserModel.HSSFWorkbook exBook = new NPOI.HSSF.UserModel.HSSFWorkbook();
            //第二步：创建Excel对象的工作簿
            NPOI.SS.UserModel.ISheet sheet = exBook.CreateSheet();
            //第三步：Excel表头设置
            //给sheet添加第一行的头部标题
            NPOI.SS.UserModel.IRow row1 = sheet.CreateRow(0);//创建行
            row1.CreateCell(0).SetCellValue("运输路线");
            row1.CreateCell(1).SetCellValue("费用项目");
            row1.CreateCell(2).SetCellValue("箱型");
            row1.CreateCell(3).SetCellValue("报关方式");
            row1.CreateCell(4).SetCellValue("金额");
            row1.CreateCell(5).SetCellValue("币种");
            row1.CreateCell(6).SetCellValue("报价人");
            row1.CreateCell(7).SetCellValue("备注");
            //4、循环写入数据
            for (var i = 0; i < listOfferDetail.Count(); i++)
            {
                NPOI.SS.UserModel.IRow rowTemp = sheet.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listOfferDetail[i].HaulWayDescription);
                rowTemp.CreateCell(1).SetCellValue(listOfferDetail[i].ExpenseName);
                rowTemp.CreateCell(2).SetCellValue(listOfferDetail[i].CabinetType);
                rowTemp.CreateCell(3).SetCellValue(listOfferDetail[i].EtryClasses);
                rowTemp.CreateCell(4).SetCellValue(listOfferDetail[i].Money.ToString());
                rowTemp.CreateCell(5).SetCellValue(listOfferDetail[i].Currency);
                rowTemp.CreateCell(6).SetCellValue(listOfferDetail[i].StaffName);
                rowTemp.CreateCell(7).SetCellValue(listOfferDetail[i].Remark);
            }
            //6、文件名
            var fileName = "广东信息科技有限公司客户标准运费报价明细报表" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            //5、将Excel表格转化为文件流输出
            MemoryStream exStream = new MemoryStream();
            exBook.Write(exStream);
            //输出之前调用Seek（偏移量，游标位置）方法：确定流开始的位置
            exStream.Seek(0, SeekOrigin.Begin);

            return File(exStream, "application/vnd.ms-excel", fileName);
        }

        /// <summary>
        /// 导出客户应收运费上方表格报价表
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportReceivableOffer(int? OfferID, string myArray)
        {
            var listOffer = (from tbOffer in myModels.SYS_Offer
                             join tbClient in myModels.SYS_Client on tbOffer.ClientID equals tbClient.ClientID
                             where tbOffer.OfferType.Trim() == "客户应收费用"
                             orderby tbOffer.OfferID descending
                             select new ClientVo
                             {
                                 OfferID = tbOffer.OfferID,
                                 OfferDate1 = tbOffer.OfferDate.ToString(),
                                 OfferDate = tbOffer.OfferDate,
                                 TakeEffectDate1 = tbOffer.TakeEffectDate.ToString(),
                                 LoseEfficacyDate1 = tbOffer.LoseEfficacyDate.ToString(),
                                 WhetherShuii = "",
                                 WhetherShui = tbOffer.WhetherShui,
                                 Remark = tbOffer.Remark,
                                 ClientAbbreviation = tbClient.ClientAbbreviation,
                                 ClientCode = tbClient.ClientCode,
                                 ClientID = tbClient.ClientID,

                             }).ToList();
            for (int i = 0; i < listOffer.Count(); i++)  //转换true false 为是否
            {
                if (listOffer[i].WhetherShui == true)
                {
                    listOffer[i].WhetherShuii = "是";
                }
                else
                {
                    listOffer[i].WhetherShuii = "否";
                }
            }

            if (!string.IsNullOrEmpty(myArray))
            {
                string[] AAA = myArray.Split(',');//用英文状态下的逗号分割字符串
                for (int i = 0; i < listOffer.Count(); i++)
                {
                    var F = listOffer[i].OfferID.ToString();
                    int id = Array.IndexOf(AAA, F);
                    if (id == -1)
                    {
                        listOffer.Remove(listOffer[i]);
                        i = 0;
                    }
                }
            }

            //Excel表格的创建步骤
            //第一步：创建Excel对象
            NPOI.HSSF.UserModel.HSSFWorkbook exBook = new NPOI.HSSF.UserModel.HSSFWorkbook();
            //第二步：创建Excel对象的工作簿
            NPOI.SS.UserModel.ISheet sheet = exBook.CreateSheet();
            //第三步：Excel表头设置
            //给sheet添加第一行的头部标题
            NPOI.SS.UserModel.IRow row1 = sheet.CreateRow(0);//创建行
            row1.CreateCell(0).SetCellValue("客户名称");
            row1.CreateCell(1).SetCellValue("客户代码");
            row1.CreateCell(2).SetCellValue("报价日期");
            row1.CreateCell(3).SetCellValue("生效日期");
            row1.CreateCell(4).SetCellValue("失效日期");
            row1.CreateCell(5).SetCellValue("是否含税");
            row1.CreateCell(6).SetCellValue("备注");
            //4、循环写入数据
            for (var i = 0; i < listOffer.Count(); i++)
            {
                NPOI.SS.UserModel.IRow rowTemp = sheet.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listOffer[i].ClientAbbreviation);
                rowTemp.CreateCell(1).SetCellValue(listOffer[i].ClientCode);
                rowTemp.CreateCell(2).SetCellValue(listOffer[i].OfferDate1);
                rowTemp.CreateCell(3).SetCellValue(listOffer[i].TakeEffectDate1);
                rowTemp.CreateCell(4).SetCellValue(listOffer[i].LoseEfficacyDate1);
                rowTemp.CreateCell(5).SetCellValue(listOffer[i].WhetherShuii);
                rowTemp.CreateCell(6).SetCellValue(listOffer[i].Remark);
            }
            //6、文件名
            var fileName = "广东信息科技有限公司客户应收运费报价报表" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            //5、将Excel表格转化为文件流输出
            MemoryStream exStream = new MemoryStream();
            exBook.Write(exStream);
            //输出之前调用Seek（偏移量，游标位置）方法：确定流开始的位置
            exStream.Seek(0, SeekOrigin.Begin);

            return File(exStream, "application/vnd.ms-excel", fileName);
        }

        /// <summary>
        /// 导出客户应收运费下方表格的报价明细表
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportReceivableOfferDetail(int? OfferID, string myArray)
        {
            var listOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                   join tbExpense in myModels.SYS_Expense on tbOfferDetail.ExpenseID equals tbExpense.ExpenseID
                                   join tbHaulWay in myModels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWay.HaulWayID
                                   join tbMention in myModels.SYS_Mention on tbHaulWay.MentionAreaID equals tbMention.MentionID
                                   join tbLoading in myModels.SYS_Gatedot on tbHaulWay.LoadingAreaID equals tbLoading.GatedotID
                                   join tbAlsoTank in myModels.SYS_Mention on tbHaulWay.AlsoTankAreaID equals tbAlsoTank.MentionID
                                   join tbStaff in myModels.SYS_Staff on tbOfferDetail.StaffID equals tbStaff.StaffID
                                   orderby tbOfferDetail.OfferDetailID descending
                                   select new ClientVo
                                   {
                                       HaulWayID = tbHaulWay.HaulWayID,
                                       OfferID = tbOfferDetail.OfferID,
                                       OfferDetailID = tbOfferDetail.OfferDetailID,
                                       Abbreviation = tbMention.Abbreviation.Trim(),
                                       Abbreviationn = tbAlsoTank.Abbreviation.Trim(),
                                       GatedotName = tbLoading.GatedotName.Trim(),
                                       HaulWayDescription = tbMention.Abbreviation.Trim() + "-" + tbLoading.GatedotName.Trim() + "-" + tbAlsoTank.Abbreviation.Trim(),
                                       ExpenseName = tbExpense.ExpenseName,
                                       CabinetType = tbOfferDetail.CabinetType.Trim(),
                                       EtryClasses = tbOfferDetail.EtryClasses.Trim(),
                                       Money = tbOfferDetail.Money,
                                       Currency = tbOfferDetail.Currency,
                                       Remark = tbOfferDetail.Remark,
                                       StaffName = tbStaff.StaffName,
                                       BoxQuantity = tbOfferDetail.BoxQuantity,
                                   }).ToList();

            if (OfferID > 0)
            {
                listOfferDetail = listOfferDetail.Where(m => m.OfferID == OfferID).ToList();
            }

            if (!string.IsNullOrEmpty(myArray))
            {
                string[] AAA = myArray.Split(',');//用英文状态下的逗号分割字符串
                for (int i = 0; i < listOfferDetail.Count(); i++)
                {
                    var F = listOfferDetail[i].OfferDetailID.ToString();
                    int id = Array.IndexOf(AAA, F);
                    if (id == -1)
                    {
                        listOfferDetail.Remove(listOfferDetail[i]);
                        i = 0;
                    }
                }
            }

            //Excel表格的创建步骤
            //第一步：创建Excel对象
            NPOI.HSSF.UserModel.HSSFWorkbook exBook = new NPOI.HSSF.UserModel.HSSFWorkbook();
            //第二步：创建Excel对象的工作簿
            NPOI.SS.UserModel.ISheet sheet = exBook.CreateSheet();
            //第三步：Excel表头设置
            //给sheet添加第一行的头部标题
            NPOI.SS.UserModel.IRow row1 = sheet.CreateRow(0);//创建行
            row1.CreateCell(0).SetCellValue("运输路线");
            row1.CreateCell(1).SetCellValue("费用项目");
            row1.CreateCell(2).SetCellValue("箱型");
            row1.CreateCell(3).SetCellValue("箱量");
            row1.CreateCell(4).SetCellValue("报关方式");
            row1.CreateCell(5).SetCellValue("金额");
            row1.CreateCell(6).SetCellValue("币种");
            row1.CreateCell(7).SetCellValue("报价人");
            //4、循环写入数据
            for (var i = 0; i < listOfferDetail.Count(); i++)
            {
                NPOI.SS.UserModel.IRow rowTemp = sheet.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listOfferDetail[i].HaulWayDescription);
                rowTemp.CreateCell(1).SetCellValue(listOfferDetail[i].ExpenseName);
                rowTemp.CreateCell(2).SetCellValue(listOfferDetail[i].CabinetType);
                rowTemp.CreateCell(3).SetCellValue(listOfferDetail[i].BoxQuantity);
                rowTemp.CreateCell(4).SetCellValue(listOfferDetail[i].EtryClasses);
                rowTemp.CreateCell(5).SetCellValue(listOfferDetail[i].Money.ToString());
                rowTemp.CreateCell(6).SetCellValue(listOfferDetail[i].Currency);
                rowTemp.CreateCell(7).SetCellValue(listOfferDetail[i].StaffName);
            }
            //6、文件名
            var fileName = "广东信息科技有限公司客户应收运费报价明细报表" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            //5、将Excel表格转化为文件流输出
            MemoryStream exStream = new MemoryStream();
            exBook.Write(exStream);
            //输出之前调用Seek（偏移量，游标位置）方法：确定流开始的位置
            exStream.Seek(0, SeekOrigin.Begin);

            return File(exStream, "application/vnd.ms-excel", fileName);
        }

        /// <summary>
        /// 导出车队标准运费上方表格的报价
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportMotorcadeOffce(int? ClientID, int? OfferID, string ClientAbbreviation, string OfferDate, string OfferDatee, bool? WhetherShui, string myArray)
        {
            var listOffer = (from tbOffer in myModels.SYS_Offer
                             join tbClient in myModels.SYS_Client on tbOffer.ClientID equals tbClient.ClientID
                             where tbOffer.OfferType.Trim() == "车队标准运费"
                             orderby tbOffer.OfferID descending
                             select new ClientVo
                             {
                                 OfferID = tbOffer.OfferID,
                                 OfferDate1 = tbOffer.OfferDate.ToString(),
                                 OfferDate = tbOffer.OfferDate,
                                 TakeEffectDate1 = tbOffer.TakeEffectDate.ToString(),
                                 LoseEfficacyDate1 = tbOffer.LoseEfficacyDate.ToString(),
                                 WhetherShuii = "",
                                 WhetherShui = tbOffer.WhetherShui,
                                 Remark = tbOffer.Remark,
                                 ClientAbbreviation = tbClient.ClientAbbreviation,
                                 ClientCode = tbClient.ClientCode,
                                 ClientID = tbClient.ClientID,

                             }).ToList();

            for (int i = 0; i < listOffer.Count(); i++)  //转换true false 为是否
            {
                if (listOffer[i].WhetherShui == true)
                {
                    listOffer[i].WhetherShuii = "是";
                }
                else
                {
                    listOffer[i].WhetherShuii = "否";
                }
            }

            if (ClientID > 0)
            {
                listOffer = listOffer.Where(m => m.ClientID == ClientID).ToList();
            }
            if (!string.IsNullOrEmpty(ClientAbbreviation))//客户简称
            {
                ClientAbbreviation = ClientAbbreviation.Trim();
                listOffer = listOffer.Where(p => p.ClientAbbreviation.Contains(ClientAbbreviation)).ToList();
            }

            //判断时间段去查询的方法，有三种情况，所以需要写三个判断
            if (!string.IsNullOrEmpty(OfferDate) && string.IsNullOrEmpty(OfferDatee))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(OfferDate);
                    listOffer = listOffer.Where(m => m.OfferDate == Time).ToList();
                }
                catch (Exception)
                {
                    listOffer = listOffer.Where(m => m.OfferID == OfferID).ToList();
                }
            }
            else if (string.IsNullOrEmpty(OfferDate) && !string.IsNullOrEmpty(OfferDatee))
            {
                try
                {
                    DateTime Times = Convert.ToDateTime(OfferDatee);
                    listOffer = listOffer.Where(m => m.OfferDate <= Times).ToList();
                }
                catch (Exception)
                {
                    listOffer = listOffer.Where(m => m.OfferID == OfferID).ToList();
                }
            }
            else if (!string.IsNullOrEmpty(OfferDate) && !string.IsNullOrEmpty(OfferDatee))
            {
                try
                {
                    DateTime Time = Convert.ToDateTime(OfferDate);
                    DateTime Times = Convert.ToDateTime(OfferDatee);
                    listOffer = listOffer.Where(m => m.OfferDate >= Time && m.OfferDate <= Times).ToList();
                }
                catch (Exception)
                {
                    listOffer = listOffer.Where(m => m.OfferID == OfferID).ToList();
                }
            }

            if (WhetherShui != null)//是否含税  只要判断不为空就可以了
            {
                listOffer = listOffer.Where(m => m.WhetherShui == WhetherShui).ToList();
            }

            if (!string.IsNullOrEmpty(myArray))
            {
                string[] AAA = myArray.Split(',');//用英文状态下的逗号分割字符串
                for (int i = 0; i < listOffer.Count(); i++)
                {
                    var F = listOffer[i].OfferID.ToString();
                    int id = Array.IndexOf(AAA, F);
                    if (id == -1)
                    {
                        listOffer.Remove(listOffer[i]);
                        i = 0;
                    }
                }
            }

            //Excel表格的创建步骤
            //第一步：创建Excel对象
            NPOI.HSSF.UserModel.HSSFWorkbook exBook = new NPOI.HSSF.UserModel.HSSFWorkbook();
            //第二步：创建Excel对象的工作簿
            NPOI.SS.UserModel.ISheet sheet = exBook.CreateSheet();
            //第三步：Excel表头设置
            //给sheet添加第一行的头部标题
            NPOI.SS.UserModel.IRow row1 = sheet.CreateRow(0);//创建行
            row1.CreateCell(0).SetCellValue("客户名称");
            row1.CreateCell(1).SetCellValue("客户代码");
            row1.CreateCell(2).SetCellValue("报价日期");
            row1.CreateCell(3).SetCellValue("生效日期");
            row1.CreateCell(4).SetCellValue("失效日期");
            row1.CreateCell(5).SetCellValue("备注");
            row1.CreateCell(6).SetCellValue("是否含税");
            //4、循环写入数据
            for (var i = 0; i < listOffer.Count(); i++)
            {
                NPOI.SS.UserModel.IRow rowTemp = sheet.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listOffer[i].ClientAbbreviation);
                rowTemp.CreateCell(1).SetCellValue(listOffer[i].ClientCode);
                rowTemp.CreateCell(2).SetCellValue(listOffer[i].OfferDate1);
                rowTemp.CreateCell(3).SetCellValue(listOffer[i].TakeEffectDate1);
                rowTemp.CreateCell(4).SetCellValue(listOffer[i].LoseEfficacyDate1);
                rowTemp.CreateCell(5).SetCellValue(listOffer[i].Remark);
                rowTemp.CreateCell(6).SetCellValue(listOffer[i].WhetherShuii);
            }
            //6、文件名
            var fileName = "广东信息科技有限公司车队标准运费报价报表" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            //5、将Excel表格转化为文件流输出
            MemoryStream exStream = new MemoryStream();
            exBook.Write(exStream);
            //输出之前调用Seek（偏移量，游标位置）方法：确定流开始的位置
            exStream.Seek(0, SeekOrigin.Begin);

            return File(exStream, "application/vnd.ms-excel", fileName);
        }


        /// <summary>
        /// 导出下方车队标准运费报价明细的表格
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportMotorcadeOfferDetail(int? OfferID, string myArray)
        {
            var listOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                   join tbExpense in myModels.SYS_Expense on tbOfferDetail.ExpenseID equals tbExpense.ExpenseID
                                   join tbHaulWay in myModels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWay.HaulWayID
                                   join tbMention in myModels.SYS_Mention on tbHaulWay.MentionAreaID equals tbMention.MentionID
                                   join tbLoading in myModels.SYS_Gatedot on tbHaulWay.LoadingAreaID equals tbLoading.GatedotID
                                   join tbAlsoTank in myModels.SYS_Mention on tbHaulWay.AlsoTankAreaID equals tbAlsoTank.MentionID
                                   join tbStaff in myModels.SYS_Staff on tbOfferDetail.StaffID equals tbStaff.StaffID
                                   orderby tbOfferDetail.OfferDetailID descending
                                   select new ClientVo
                                   {
                                       HaulWayID = tbHaulWay.HaulWayID,
                                       OfferID = tbOfferDetail.OfferID,
                                       OfferDetailID = tbOfferDetail.OfferDetailID,
                                       Abbreviation = tbMention.Abbreviation.Trim(),
                                       Abbreviationn = tbAlsoTank.Abbreviation.Trim(),
                                       GatedotName = tbLoading.GatedotName.Trim(),
                                       HaulWayDescription = tbMention.Abbreviation.Trim() + "-" + tbLoading.GatedotName.Trim() + "-" + tbAlsoTank.Abbreviation.Trim(),
                                       ExpenseName = tbExpense.ExpenseName,
                                       CabinetType = tbOfferDetail.CabinetType.Trim(),
                                       EtryClasses = tbOfferDetail.EtryClasses.Trim(),
                                       Money = tbOfferDetail.Money,
                                       Currency = tbOfferDetail.Currency,
                                       Remark = tbOfferDetail.Remark,
                                       StaffName = tbStaff.StaffName,
                                       BoxQuantity = tbOfferDetail.BoxQuantity,
                                   }).ToList();

            if (OfferID > 0)
            {
                listOfferDetail = listOfferDetail.Where(m => m.OfferID == OfferID).ToList();
            }

            if (!string.IsNullOrEmpty(myArray))
            {
                string[] AAA = myArray.Split(',');//用英文状态下的逗号分割字符串
                for (int i = 0; i < listOfferDetail.Count(); i++)
                {
                    var F = listOfferDetail[i].OfferDetailID.ToString();
                    int id = Array.IndexOf(AAA, F);
                    if (id == -1)
                    {
                        listOfferDetail.Remove(listOfferDetail[i]);
                        i = 0;
                    }
                }
            }

            //Excel表格的创建步骤
            //第一步：创建Excel对象
            NPOI.HSSF.UserModel.HSSFWorkbook exBook = new NPOI.HSSF.UserModel.HSSFWorkbook();
            //第二步：创建Excel对象的工作簿
            NPOI.SS.UserModel.ISheet sheet = exBook.CreateSheet();
            //第三步：Excel表头设置
            //给sheet添加第一行的头部标题
            NPOI.SS.UserModel.IRow row1 = sheet.CreateRow(0);//创建行
            row1.CreateCell(0).SetCellValue("运输路线");
            row1.CreateCell(1).SetCellValue("费用项目");
            row1.CreateCell(2).SetCellValue("箱型");
            row1.CreateCell(3).SetCellValue("箱量");
            row1.CreateCell(4).SetCellValue("报关方式");
            row1.CreateCell(5).SetCellValue("金额");
            row1.CreateCell(6).SetCellValue("币种");
            row1.CreateCell(7).SetCellValue("报价人");
            //4、循环写入数据
            for (var i = 0; i < listOfferDetail.Count(); i++)
            {
                NPOI.SS.UserModel.IRow rowTemp = sheet.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listOfferDetail[i].HaulWayDescription);
                rowTemp.CreateCell(1).SetCellValue(listOfferDetail[i].ExpenseName);
                rowTemp.CreateCell(2).SetCellValue(listOfferDetail[i].CabinetType);
                rowTemp.CreateCell(3).SetCellValue(listOfferDetail[i].BoxQuantity);
                rowTemp.CreateCell(4).SetCellValue(listOfferDetail[i].EtryClasses);
                rowTemp.CreateCell(5).SetCellValue(listOfferDetail[i].Money.ToString());
                rowTemp.CreateCell(6).SetCellValue(listOfferDetail[i].Currency);
                rowTemp.CreateCell(7).SetCellValue(listOfferDetail[i].StaffName);
            }
            //6、文件名
            var fileName = "广东信息科技有限公司车队标准运费报价明细报表" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            //5、将Excel表格转化为文件流输出
            MemoryStream exStream = new MemoryStream();
            exBook.Write(exStream);
            //输出之前调用Seek（偏移量，游标位置）方法：确定流开始的位置
            exStream.Seek(0, SeekOrigin.Begin);

            return File(exStream, "application/vnd.ms-excel", fileName);
        }


        /// <summary>
        /// 导出司机产值下方表格的报价明细表格数据
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportChauffeurProduce(int? OfferID, string myArray)
        {
            var listOfferDetail = (from tbOfferDetail in myModels.SYS_OfferDetail
                                   join tbExpense in myModels.SYS_Expense on tbOfferDetail.ExpenseID equals tbExpense.ExpenseID
                                   join tbHaulWay in myModels.SYS_HaulWay on tbOfferDetail.HaulWayID equals tbHaulWay.HaulWayID
                                   join tbMention in myModels.SYS_Mention on tbHaulWay.MentionAreaID equals tbMention.MentionID
                                   join tbLoading in myModels.SYS_Gatedot on tbHaulWay.LoadingAreaID equals tbLoading.GatedotID
                                   join tbAlsoTank in myModels.SYS_Mention on tbHaulWay.AlsoTankAreaID equals tbAlsoTank.MentionID
                                   join tbStaff in myModels.SYS_Staff on tbOfferDetail.StaffID equals tbStaff.StaffID
                                   orderby tbOfferDetail.OfferDetailID descending
                                   select new ClientVo
                                   {
                                       HaulWayID = tbHaulWay.HaulWayID,
                                       OfferID = tbOfferDetail.OfferID,
                                       OfferDetailID = tbOfferDetail.OfferDetailID,
                                       Abbreviation = tbMention.Abbreviation.Trim(),
                                       Abbreviationn = tbAlsoTank.Abbreviation.Trim(),
                                       GatedotName = tbLoading.GatedotName.Trim(),
                                       HaulWayDescription = tbMention.Abbreviation + "-" + tbLoading.GatedotName + "-" + tbAlsoTank.Abbreviation,
                                       ExpenseName = tbExpense.ExpenseName,
                                       CabinetType = tbOfferDetail.CabinetType.Trim(),
                                       EtryClasses = tbOfferDetail.EtryClasses.Trim(),
                                       Money = tbOfferDetail.Money,
                                       Currency = tbOfferDetail.Currency,
                                       Remark = tbOfferDetail.Remark,
                                       StaffName = tbStaff.StaffName,
                                       BoxQuantity = tbOfferDetail.BoxQuantity
                                   }).ToList();

            if (OfferID > 0)
            {
                listOfferDetail = listOfferDetail.Where(m => m.OfferID == OfferID).ToList();
            }
            if (!string.IsNullOrEmpty(myArray))
            {
                string[] AAA = myArray.Split(',');//用英文状态下的逗号分割字符串
                for (int i = 0; i < listOfferDetail.Count(); i++)
                {
                    var F = listOfferDetail[i].OfferID.ToString();
                    int id = Array.IndexOf(AAA, F);
                    if (id == -1)
                    {
                        listOfferDetail.Remove(listOfferDetail[i]);
                        i = 0;
                    }
                }
            }

            //Excel表格的创建步骤
            //第一步：创建Excel对象
            NPOI.HSSF.UserModel.HSSFWorkbook exBook = new NPOI.HSSF.UserModel.HSSFWorkbook();
            //第二步：创建Excel对象的工作簿
            NPOI.SS.UserModel.ISheet sheet = exBook.CreateSheet();
            //第三步：Excel表头设置
            //给sheet添加第一行的头部标题
            NPOI.SS.UserModel.IRow row1 = sheet.CreateRow(0);//创建行
            row1.CreateCell(0).SetCellValue("运输路线");
            row1.CreateCell(1).SetCellValue("费用项目");
            row1.CreateCell(2).SetCellValue("箱型");
            row1.CreateCell(3).SetCellValue("箱量");
            row1.CreateCell(4).SetCellValue("报关方式");
            row1.CreateCell(5).SetCellValue("金额");
            row1.CreateCell(6).SetCellValue("币种");
            row1.CreateCell(7).SetCellValue("报价人");
            row1.CreateCell(8).SetCellValue("备注");
            //4、循环写入数据
            for (var i = 0; i < listOfferDetail.Count(); i++)
            {
                NPOI.SS.UserModel.IRow rowTemp = sheet.CreateRow(i + 1);
                rowTemp.CreateCell(0).SetCellValue(listOfferDetail[i].HaulWayDescription);
                rowTemp.CreateCell(1).SetCellValue(listOfferDetail[i].ExpenseName);
                rowTemp.CreateCell(2).SetCellValue(listOfferDetail[i].CabinetType);
                rowTemp.CreateCell(3).SetCellValue(listOfferDetail[i].BoxQuantity);
                rowTemp.CreateCell(4).SetCellValue(listOfferDetail[i].EtryClasses);
                rowTemp.CreateCell(5).SetCellValue(listOfferDetail[i].Money.ToString());
                rowTemp.CreateCell(6).SetCellValue(listOfferDetail[i].Currency);
                rowTemp.CreateCell(7).SetCellValue(listOfferDetail[i].StaffName);
                rowTemp.CreateCell(8).SetCellValue(listOfferDetail[i].Remark);
            }
            //6、文件名
            var fileName = "广东信息科技有限公司司机产值报价明细报表" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            //5、将Excel表格转化为文件流输出
            MemoryStream exStream = new MemoryStream();
            exBook.Write(exStream);
            //输出之前调用Seek（偏移量，游标位置）方法：确定流开始的位置
            exStream.Seek(0, SeekOrigin.Begin);

            return File(exStream, "application/vnd.ms-excel", fileName);
        }



        #endregion

        #region 查询下拉框数据

        /// <summary>
        /// 绑定门点下拉框数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectGatedot()
        {
            var listGatedot = (from tbGatedot in myModels.SYS_Gatedot
                               select new SelectVo
                               {
                                   id = tbGatedot.GatedotID,
                                   text = tbGatedot.GatedotName,
                               }).ToList();
            listGatedot = Common.Tools.SetSelectJson(listGatedot);
            return Json(listGatedot, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询报价人下拉框
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectStaff()
        {
            var listStaff = (from tbStaff in myModels.SYS_Staff
                             select new SelectVo
                             {
                                 id = tbStaff.StaffID,
                                 text = tbStaff.StaffName
                             }).ToList();
            listStaff = Common.Tools.SetSelectJson(listStaff);
            return Json(listStaff, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询费用项目名称下拉框数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectExpense()
        {
            var listExpense = (from tbExpense in myModels.SYS_Expense
                               select new SelectVo
                               {
                                   id = tbExpense.ExpenseID,
                                   text = tbExpense.ExpenseName,
                               }).ToList();
            listExpense = Common.Tools.SetSelectJson(listExpense);
            return Json(listExpense, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 绑定运输路线下拉框数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectHaulWay()
        {
            var listHaulWay = (from tbHaulWay in myModels.SYS_HaulWay
                               join tbMention in myModels.SYS_Mention on tbHaulWay.MentionAreaID equals tbMention.MentionID
                               join tbLoading in myModels.SYS_Gatedot on tbHaulWay.LoadingAreaID equals tbLoading.GatedotID
                               join tbAlsoTank in myModels.SYS_Mention on tbHaulWay.AlsoTankAreaID equals tbAlsoTank.MentionID
                               select new SelectVo
                               {
                                   id = tbHaulWay.HaulWayID,
                                   text = tbMention.Abbreviation + "-" + tbLoading.GatedotName + "-" + tbAlsoTank.Abbreviation
                               }).ToList();
            listHaulWay = Common.Tools.SetSelectJson(listHaulWay);
            return Json(listHaulWay, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 绑定客户代码下拉框数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectClientCode()
        {
            var listCode = (from tbCode in myModels.SYS_Client
                            select new SelectVo
                            {
                                id = tbCode.ClientID,
                                text = tbCode.ClientCode,
                            }).ToList();
            listCode = Common.Tools.SetSelectJson(listCode);
            return Json(listCode, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询树形图的表格数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectTree()
        {
            var listTree = (from tbTree in myModels.SYS_Offer
                            join tbClient in myModels.SYS_Client on tbTree.ClientID equals tbClient.ClientID
                            where tbTree.OfferType.Trim() == "客户应收费用"
                            select new 
                            {
                                id = tbClient.ClientID,
                                name = tbClient.ClientAbbreviation,
                            }).Distinct().ToList();

            return Json(listTree, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询树形图的表格数据
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectTree1()
        {
            var listTree = (from tbTree in myModels.SYS_Offer
                            join tbClient in myModels.SYS_Client on tbTree.ClientID equals tbClient.ClientID
                            where tbTree.OfferType.Trim() == "车队标准运费"
                            select new
                            {
                                id = tbClient.ClientID,
                                name = tbClient.ClientAbbreviation,
                            }).Distinct().ToList();

            return Json(listTree, JsonRequestBehavior.AllowGet);
        }





        #endregion
    }
}
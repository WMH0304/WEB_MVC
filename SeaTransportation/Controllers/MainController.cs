using SeaTransportation.Models;
using SeaTransportation.Vo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SeaTransportation.Controllers
{
    public class MainController : Controller
    {
        Models.SeaTransportationEntities mymodels = new Models.SeaTransportationEntities();
        // GET: Main
        public ActionResult Login()
        {
            return View();
        }
        public ActionResult Main()
        {
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());//将Session中的用户ID的字符串强制装换为整型,赋值给UserID
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");  //如果代码出错就放回到登录页面(即Session["UserID"]为空,代码会出错)
            }
            return View();
        }
        //登录
        public ActionResult DengLu(string User,string Password,int MessageID,int DepartmentID)
        {
            string Data = "";
            try
            {
                var user = (from tbUser in mymodels.SYS_User
                            join tbStaff in mymodels.SYS_Staff on tbUser.StaffID equals tbStaff.StaffID
                            join tbDepartment in mymodels.SYS_Department on tbStaff.DepartmentID equals tbDepartment.DepartmentID
                            join tbMessage in mymodels.SYS_Message on tbDepartment.MessageID equals tbMessage.MessageID
                            where tbUser.AccountName == User
                            select new
                            {
                                StaffName = tbStaff.StaffName,
                                StaffID = tbStaff.StaffID,
                                Password = tbUser.password.Trim(),
                                MessageID = tbMessage.MessageID,
                                ChineseName = tbMessage.ChineseName,
                                DepartmentID = tbDepartment.DepartmentID,
                                DepartmentName = tbDepartment.DepartmentName,
                                UserID = tbUser.UserID
                            }).ToList();
                if (user.Count() > 0)
                {
                    if (user[0].Password == Password)
                    {
                        if (user[0].MessageID == MessageID)
                        {
                            if (user[0].DepartmentID == DepartmentID)
                            {
                                Data = "登录成功";
                                Session["AccountName"] = User;
                                Session["ChineseName"] = user[0].ChineseName;
                                Session["StaffID"] = user[0].StaffID;
                                Session["StaffName"] = user[0].StaffName;
                                Session["DepartmentName"] = user[0].DepartmentName;
                                Session["UserID"] = user[0].UserID;
                            }
                            else
                            {
                                Data = "部门错误";
                            }
                        }
                        else
                        {
                            Data = "公司错误";
                        }
                    }
                    else
                    {
                        Data = "密码错误";
                    }
                }
                else
                {
                    Data = "不存在此用户";
                }
                return Json(Data, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {

                return null;
            }
        }
        public ActionResult Home()
        {
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());
                List<SYS_User> user = mymodels.SYS_User.Where(m => m.UserID == UserID).ToList();
                user[0].FinallyRegisterTime = DateTime.Now;
                mymodels.Entry(user[0]).State = System.Data.Entity.EntityState.Modified;
                mymodels.SaveChanges();
                string[] main = new string[4] { Session["AccountName"].ToString(), Session["ChineseName"].ToString(), Session["DepartmentName"].ToString(), DateTime.Now.ToString() };
                return Json(main, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {

                return null;
            }
        }
        //主页面
        public ActionResult index()
        {
            try
            {
                int UserID = Convert.ToInt32(Session["UserID"].ToString());//将Session中的用户ID的字符串强制装换为整型,赋值给UserID
            }
            catch (Exception)
            {
                return Redirect("/Main/Login");  //如果代码出错就放回到登录页面(即Session["UserID"]为空,代码会出错)
            }
            return View();
        }
        //公司下拉框
        public ActionResult Message()
        {
            try
            {
                var Message = from tbMessage in mymodels.SYS_Message
                              select new
                              {
                                  id = tbMessage.MessageID,
                                  name = tbMessage.ChineseName
                              };
                return Json(Message, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {

                return null;
            }
        }
        //部门下拉框
        public ActionResult Department(short? MessageID)
        {
            try
            {
                var Department = from tbDepartment in mymodels.SYS_Department
                                 where tbDepartment.MessageID == MessageID
                                 select new
                                 {
                                     id = tbDepartment.DepartmentID,
                                     name = tbDepartment.DepartmentName
                                 };
                return Json(Department, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {

                return null;
            }
        }
        public ActionResult Tongji() {
            var data = from tbCharge in mymodels.SYS_Charge
                        join tbExpense in mymodels.SYS_Expense on tbCharge.ExpenseID equals tbExpense.ExpenseID
                       select new
                             {
                                  tbExpense.SettleAccounts,
                                 tbCharge.UnitPrice
                             };
            EtrustVo Etrust = new EtrustVo();
            Etrust.S = data.Where(m => m.SettleAccounts.Trim() == "应收").Sum(m => m.UnitPrice + m.UnitPrice);
            Etrust.C = data.Where(m => m.SettleAccounts.Trim() == "应付").Sum(m => m.UnitPrice + m.UnitPrice);
            Etrust.M = mymodels.SYS_Etrust.Select(m => m.EtrustID).Count();
            Etrust.J = mymodels.SYS_Client.Select(m => m.ClientID).Count();
            return Json(Etrust, JsonRequestBehavior.AllowGet);
        }

        //public ActionResult SelectStatisticsData()
        //{
        //    try
        //    {
        //        var ArchitecturalNature = (from tbCharge in mymodels.SYS_Charge
        //                                  join tbExpense in mymodels.SYS_Expense on tbCharge.ExpenseID equals tbExpense.ExpenseID
        //                                  select new
        //                                  {
        //                                      tbExpense.SettleAccounts,
        //                                      tbCharge.UnitPrice
        //                                  }).ToList();
        //        int totleCount = ArchitecturalNature.Count() == 0 ? 1 : ArchitecturalNature.Count();
        //        EtrustVo listStatistics = new EtrustVo();
        //        listStatistics.S = totleCount;
        //        listStatistics.J = ArchitecturalNature.Count(m => m.SettleAccounts == "应收");
        //        listStatistics.M = listStatistics.J / totleCount * 100;
        //        listStatistics.C = ArchitecturalNature.Count(m => m.SettleAccounts == "应收");
        //        listStatistics.UnitPrice = listStatistics.C / totleCount * 100;
        //        return Json(listStatistics, JsonRequestBehavior.AllowGet);

        //    }
        //    catch (Exception)
        //    {
        //        return Json("数据异常", JsonRequestBehavior.AllowGet);
        //    }
        //}
    }
}
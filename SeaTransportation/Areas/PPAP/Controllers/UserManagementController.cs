using Hotel.VO;
using SeaTransportation.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Web.Mvc;

namespace SeaTransportation.Areas.PPAP.Controllers
{
    public class UserManagementController : Controller
    {
        Models.SeaTransportationEntities myModels = new Models.SeaTransportationEntities();
        // GET: PPAP/UserManagement
        public ActionResult UserManagement()
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

        public ActionResult Staff(int? DepartmentID)
        {
            var List = from tbStaff in myModels.SYS_Staff
                       where tbStaff.DepartmentID == DepartmentID
                       select new
                       {
                           id = tbStaff.StaffID,
                           name = tbStaff.StaffName,
                       };
            return Json(List, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查询
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectUserManagement(BsgridPage bsgridPage, string AccountName, string StaffName, string password, int? A)
        {
            var list = (from tbUser in myModels.SYS_User
                        join tbStaff in myModels.SYS_Staff on tbUser.StaffID equals tbStaff.StaffID
                        join tbDepartment in myModels.SYS_Department on tbStaff.DepartmentID equals tbDepartment.DepartmentID
                        join tbMessage in myModels.SYS_Message on tbDepartment.MessageID equals tbMessage.MessageID
                        into tbMessage
                        from Message in tbMessage.DefaultIfEmpty()
                        select new UserManagement
                        {
                            MessageID = tbDepartment.MessageID,
                            DepartmentID = tbDepartment.DepartmentID,
                            StaffID = tbStaff.StaffID,
                            UserID = tbUser.UserID,
                            AccountName = tbUser.AccountName,             //账户名
                            password = tbUser.password,                    //密码
                            StaffName = tbStaff.StaffName,                //员工姓名
                            Time = tbUser.FinallyRegisterTime.ToString(), //最后登录时间 
                            Describe = tbUser.Describe,                   //描述
                            WhetherStart = tbUser.WhetherStart,           //是否启用
                            ChineseName = Message.ChineseName,            //所属组织
                            DepartmentName = tbDepartment.DepartmentName, //所属部门
                        }).ToList();
            #region MyRegion


            if (!string.IsNullOrEmpty(AccountName))
            {
                list = list.Where(m => m.AccountName.Contains(AccountName)).ToList();
            }
            if (!string.IsNullOrEmpty(password))
            {

                list = list.Where(m => m.password.Contains(password.Trim())).ToList();
            }
            if (!string.IsNullOrEmpty(StaffName))
            {

                list = list.Where(m => m.StaffName.Contains(StaffName)).ToList();
            }
            if (A == 1)
            {
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            #endregion
            int pageCount = list.Count();
            List<UserManagement> wcg = list.Skip(bsgridPage.GetStartIndex()).Take(bsgridPage.pageSize).ToList();
            Bsgrid<UserManagement> bsgrid = new Bsgrid<UserManagement>()
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
        /// <param name="myUser"></param>
        /// <returns></returns>
        public ActionResult InserUserManagement(SYS_User myUser)
        {
            string strMsg = "failed";


            var list = (from tbUSser in myModels.SYS_User
                        where tbUSser.AccountName == myUser.AccountName && tbUSser.Describe == myUser.Describe
                        select tbUSser).Count();
            if (list == 0)
            {
                try
                {
                    myModels.SYS_User.Add(myUser);
                    myModels.SaveChanges();
                    strMsg = "success";
                    if (myModels.SaveChanges() > 0)
                    {
                        strMsg = "新增成功！";
                    }
                }

                catch (Exception)
                {
                    strMsg = "failed";
                }
            }
            else
            {
                strMsg = "已有员工数据!";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 修改
        /// </summary>
        /// <param name="myUser"></param>
        /// <param name="UserID"></param>
        /// <returns></returns>    
        /// 修改把时间清空了，等后面获取离开时间在弄
        public ActionResult UpDateManagement(SYS_User myUser)
        {
            string strMsg = "failed";
            try
            {
                myModels.SYS_User.Add(myUser);

                myModels.Entry(myUser).State = System.Data.Entity.EntityState.Modified;
                if (myModels.SaveChanges() > 0)
                {
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
        /// 删除
        /// </summary>
        /// <param name="UserID"></param>
        /// <returns></returns>
        public ActionResult DeleteManagement(int UserID)
        {
            string strMsg = "fail";
            try
            {
                var User = myModels.SYS_User.Where(m => m.UserID == UserID).Single();
                myModels.SYS_User.Remove(User);
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
        public ActionResult FlowLeaveManagement(string AccountName, string StaffName, string FinallyRegisterTime, string Describe, string WhetherStart, string ChineseName, string DepartmentName)
        {
            var list = (from tbUser in myModels.SYS_User
                        join tbStaff in myModels.SYS_Staff on tbUser.StaffID equals tbStaff.StaffID
                        join tbDepartment in myModels.SYS_Department on tbStaff.DepartmentID equals tbDepartment.DepartmentID
                        join tbMessage in myModels.SYS_Message on tbDepartment.MessageID equals tbMessage.MessageID
                        into tbMessage
                        from Message in tbMessage.DefaultIfEmpty()
                        select new UserManagement
                        {
                            MessageID = tbDepartment.MessageID,
                            DepartmentID = tbDepartment.DepartmentID,
                            StaffID = tbStaff.StaffID,
                            UserID = tbUser.UserID,
                            AccountName = tbUser.AccountName,             //账户名
                            StaffName = tbStaff.StaffName,              //员工姓名
                            FinallyRegisterTime = tbUser.FinallyRegisterTime.ToString(),         //最后登录时间 
                            Describe = tbUser.Describe,                //描述
                            WhetherStart = tbUser.WhetherStart,            //是否启用
                            ChineseName = Message.ChineseName,            //所属组织
                            DepartmentName = tbDepartment.DepartmentName,    //所属部门
                        }).ToList();
            #region MyRegion


            if (!string.IsNullOrEmpty(AccountName))
            {
                list = list.Where(m => m.AccountName.Contains(AccountName)).ToList();
            }
            if (!string.IsNullOrEmpty(StaffName))
            {
                list = list.Where(m => m.StaffName.Contains(StaffName)).ToList();
            }
            if (!string.IsNullOrEmpty(Describe))
            {
                list = list.Where(m => m.Describe.Contains(Describe)).ToList();
            }
            if (!string.IsNullOrEmpty(ChineseName))
            {
                list = list.Where(m => m.ChineseName.Contains(ChineseName)).ToList();
            }
            #endregion
            //1、创建Excel工作簿
            NPOI.HSSF.UserModel.HSSFWorkbook leave = new NPOI.HSSF.UserModel.HSSFWorkbook();
            //2、创建工作表
            NPOI.SS.UserModel.ISheet sheet = leave.CreateSheet("葫芦娃");
            //3、设计表头
            //3.1 创建第一行
            NPOI.SS.UserModel.IRow row = sheet.CreateRow(0);
            //3.2 设计字段
            row.CreateCell(0).SetCellValue("账户名");
            row.CreateCell(1).SetCellValue("员工姓名");
            row.CreateCell(2).SetCellValue("最后登录时间");
            row.CreateCell(3).SetCellValue("描述");
            row.CreateCell(4).SetCellValue("是否启用");
            row.CreateCell(5).SetCellValue("所属组织");
            row.CreateCell(6).SetCellValue("所属部门");

            for (var i = 0; i < list.Count(); i++)
            {
                NPOI.SS.UserModel.IRow lll = sheet.CreateRow(i + 1);
                lll.CreateCell(0).SetCellValue(list[i].AccountName);
                lll.CreateCell(1).SetCellValue(list[i].StaffName);
                lll.CreateCell(2).SetCellValue(list[i].FinallyRegisterTime);
                lll.CreateCell(3).SetCellValue(list[i].Describe);
                lll.CreateCell(4).SetCellValue(list[i].WhetherStart);
                lll.CreateCell(5).SetCellValue(list[i].ChineseName);
                lll.CreateCell(6).SetCellValue(list[i].DepartmentName);
            }
            //5、将Excel表格转化为文件流输出
            MemoryStream DulaDula = new MemoryStream();
            leave.Write(DulaDula);
            //输出之前调用Seek（偏移量，游标位置）方法：确定流开始的位置
            DulaDula.Seek(0, SeekOrigin.Begin);
            //6、文件名
            string fileName = "葫芦娃" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            return File(DulaDula, "application/vnd.ms-excel", fileName);

        }
        /// <summary>
        /// 打印
        /// </summary>
        /// <returns></returns>
        /// 
        public ActionResult StampUser(string AccountName, string StaffName, string FinallyRegisterTime, string Describe, string WhetherStart, string ChineseName, string DepartmentName, int? UserID)
        {
            #region MyRegion
            var list = (from tbUser in myModels.SYS_User
                        join tbStaff in myModels.SYS_Staff on tbUser.StaffID equals tbStaff.StaffID
                        join tbDepartment in myModels.SYS_Department on tbStaff.DepartmentID equals tbDepartment.DepartmentID
                        join tbMessage in myModels.SYS_Message on tbDepartment.MessageID equals tbMessage.MessageID
                        into tbMessage
                        from Message in tbMessage.DefaultIfEmpty()
                        select new UserManagement
                        {
                            MessageID = tbDepartment.MessageID,
                            DepartmentID = tbDepartment.DepartmentID,
                            StaffID = tbStaff.StaffID,
                            UserID = tbUser.UserID,
                            AccountName = tbUser.AccountName,             //账户名
                            StaffName = tbStaff.StaffName,              //员工姓名
                            Time = tbUser.FinallyRegisterTime.ToString(),         //最后登录时间 
                            Describe = tbUser.Describe,                //描述
                            WhetherStart = tbUser.WhetherStart,            //是否启用
                            ChineseName = Message.ChineseName,            //所属组织
                            DepartmentName = tbDepartment.DepartmentName,    //所属部门
                        }).ToList();
            if (!string.IsNullOrEmpty(AccountName))
            {
                list = list.Where(m => m.AccountName.Contains(AccountName)).ToList();
            }
            if (!string.IsNullOrEmpty(StaffName))
            {
                list = list.Where(m => m.StaffName.Contains(StaffName)).ToList();
            }
            #endregion
            //将查询出来的结果转化为DataTable的格式
            DataTable dtResult = LINQToDataTable(list);
            //1、实例化数据集
            pp.ccc dbReport = new pp.ccc();
            //2、将dtResult放入数据集中名为"Message"的表格中  与数据集的表相同
            dbReport.Tables["User"].Merge(dtResult);
            //3、实例化数据报表
            pp.User rp = new pp.User();
            //4、获取报表的物理文件路径
            string strRpPath = System.Web.HttpContext.Current.Server.MapPath("~/") + "Areas\\PPAP\\pp\\User.rpt";
            //5、将报表加载到报表模板中
            rp.Load(strRpPath);
            //6、设置报表的数据源
            rp.SetDataSource(dbReport);
            //7、把报表转化为文件流输出
            Stream dbStream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
            return File(dbStream, "application/pdf");
        }

        /// <summary>
        /// 美军
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="varlist"></param>
        /// <returns></returns>
        public DataTable LINQToDataTable<T>(IEnumerable<T> varlist)
        {
            //定义要返回的DataTable对象
            DataTable dtReturn = new DataTable();
            //保存列集合的属性信息数组
            PropertyInfo[] oprops = null;
            if (varlist == null)
                return dtReturn;//安全性检查
            //循环遍历集合，使用反射获取类型的属性信息
            foreach (T rec in varlist)
            {
                //使用反射获取T类型的属性信息，返回一个PropertyInfo类型的集合
                if (oprops == null)
                {
                    oprops = ((Type)rec.GetType()).GetProperties();
                    //循环PropertyInfo数组
                    foreach (PropertyInfo pi in oprops)
                    {
                        //得到属性的类
                        Type colType = pi.PropertyType;
                        //如果属性为泛型类型
                        if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition() == typeof(Nullable<>)))
                        {
                            //获取泛型类型的参数
                            colType = colType.GetGenericArguments()[0];
                        }
                        dtReturn.Columns.Add(pi.Name, colType);
                    }
                }
                //新建一个用于添加到DataTable中的DataRow对象
                DataRow dr = dtReturn.NewRow();
                //循环遍历属性集合
                foreach (PropertyInfo pi in oprops)
                {
                    dr[pi.Name] = pi.GetValue(rec, null) == null ? DBNull.Value : pi.GetValue(rec, null);
                }
                //将具有结果值的DataRow添加到DataTable集合中
                dtReturn.Rows.Add(dr);
            }
            return dtReturn;//返回DataTable对象
        }

    }
}
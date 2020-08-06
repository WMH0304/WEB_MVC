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
    public class StaffControlController : Controller
    {
        Models.SeaTransportationEntities myModels = new Models.SeaTransportationEntities();
        // GET: PPAP/StaffControl
        /// <summary>
        /// 员工管理
        /// </summary>
        /// <returns></returns>
        public ActionResult StaffControl()
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

        #region MyRegion
        /// <summary>
        /// 查询
        /// </summary>
        /// <returns></returns>
        public ActionResult SelectStaffControl(BsgridPage bsgridPage, int? StaffID, string StaffName, string StaffNumber, string State, string Duty, string Sex, string ChineseName, string CultureLevel, string JoinJob, int? A)
        {
            var list = (from tbStaff in myModels.SYS_Staff
                        join tbDrpartment in myModels.SYS_Department on tbStaff.DepartmentID equals tbDrpartment.DepartmentID
                        join tbMessage in myModels.SYS_Message on tbDrpartment.MessageID equals tbMessage.MessageID
                        select new Staffdttcc
                        {
                            StaffID = tbStaff.StaffID,
                            DepartmentID = tbDrpartment.DepartmentID,
                            MessageID = tbMessage.MessageID,
                            StaffName = tbStaff.StaffName.Trim(),            //员工姓名
                            StaffNumber = tbStaff.StaffNumber.Trim(),        //员工编号
                            State = tbStaff.State.Trim(),                    //状态
                            Duty = tbStaff.Duty.Trim(),                      //职责
                            Sex = tbStaff.Sex.Trim(),                        //性别CultureLevel
                            ChineseName = tbMessage.ChineseName,      //所属组织
                            CultureLevel = tbStaff.CultureLevel.Trim(),       //文化程度
                            JoinJob = tbStaff.JoinJob.ToString(),      //入职日期
                            WhetherStart = tbStaff.WhetherStart,       //是否启用
                        }).ToList();
            if (!string.IsNullOrEmpty(StaffName))
            {
                list = list.Where(m => m.StaffName.Contains(StaffName)).ToList();
            }
            if (!string.IsNullOrEmpty(StaffNumber))
            {
                list = list.Where(m => m.StaffNumber.Contains(StaffNumber)).ToList();
            }
            if (!string.IsNullOrEmpty(State))
            {
                list = list.Where(m => m.State == State).ToList();
            }
            if (!string.IsNullOrEmpty(Duty))
            {
                list = list.Where(m => m.StaffNumber.Contains(Duty)).ToList();
            }
            if (!string.IsNullOrEmpty(Sex))
            {
                list = list.Where(m => m.Sex == Sex).ToList();
            }
            if (!string.IsNullOrEmpty(ChineseName))
            {
                list = list.Where(m => m.StaffNumber.Contains(ChineseName)).ToList();
            }
            if (!string.IsNullOrEmpty(CultureLevel))
            {
                list = list.Where(m => m.StaffName.Contains(CultureLevel)).ToList();
            }
            if (!string.IsNullOrEmpty(JoinJob))
            {
                list = list.Where(m => m.StaffNumber.Contains(JoinJob)).ToList();
            }

            if (StaffID > 0)
            {
                list = list.Where(m => m.StaffID == StaffID).ToList();
            }
            if (A == 1)//分流，返回列表
            {
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            int pageCount = list.Count();
            List<Staffdttcc> wcg = list.Skip(bsgridPage.GetStartIndex()).Take(bsgridPage.pageSize).ToList();
            Bsgrid<Staffdttcc> bsgrid = new Bsgrid<Staffdttcc>()
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
        public ActionResult InserStaffControl(int DepartmentID, int? StaffID, string StaffName, string StaffNumber, string State, string Duty, string Sex, string CultureLevel, string JoinJob, bool? WhetherStart)
        {
            try
            {
                //实例化一张表
                SYS_Staff myStaff = new SYS_Staff();
                //将页面写入的数据放到实例化的实体中
                myStaff.DepartmentID = Convert.ToInt32(DepartmentID);
                myStaff.StaffName = StaffName.Trim();//员工姓名\
                myStaff.StaffNumber = StaffNumber.Trim();// //员工编号
                myStaff.State = State.Trim(); //状态
                myStaff.Duty = Duty.Trim();//职责
                myStaff.Sex = Sex.Trim();////性别
                myStaff.CultureLevel = CultureLevel.Trim(); //文化程度
                myStaff.JoinJob = Convert.ToDateTime(JoinJob);        //入职日期
                myStaff.WhetherStart = Convert.ToBoolean(WhetherStart);//是否启用
                //将实例化的实体插入数据库实体
                myModels.SYS_Staff.Add(myStaff);
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
        public ActionResult UpdateStaffControl(int? StaffID, string StaffName, string StaffNumber, string State, string Duty, string Sex, string CultureLevel, string JoinJob, bool? WhetherStart)
        {
            string strMsg = "failed";
            try
            {
                Models.SYS_Staff myStaff = myModels.SYS_Staff.Where(p => p.StaffID == StaffID).Single();
                myStaff.StaffName = StaffName;//员工姓名
                myStaff.StaffNumber = StaffNumber; //员工编号
                myStaff.State = State;//状态
                myStaff.Duty = Duty;//职责
                myStaff.Sex = Sex;//性别
                myStaff.CultureLevel = CultureLevel;//文化程度
                myStaff.JoinJob = Convert.ToDateTime(JoinJob);//入职日期
                myStaff.WhetherStart = Convert.ToBoolean(WhetherStart);//是否启用\
                myModels.Entry(myStaff).State = System.Data.Entity.EntityState.Modified;
                //  DataRow
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
        /// <param name="StaffID"></param>
        /// <returns></returns>
        public ActionResult DeleteStaffControl(int StaffID)
        {
            string strMsg = "fail";
            try
            {
                var Staff = myModels.SYS_Staff.Where(m => m.StaffID == StaffID).Single();
                myModels.SYS_Staff.Remove(Staff);
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
        /// 打印
        /// </summary>
        /// <returns></returns>
        public ActionResult StampStaffControl(BsgridPage bsgridPage, int? StaffID, string StaffName, string StaffNumber, string State, string Duty, string Sex, string ChineseName, string CultureLevel, string JoinJob, int? A)
        {
            #region MyRegion
            var list = (from tbStaff in myModels.SYS_Staff
                        join tbDrpartment in myModels.SYS_Department on tbStaff.DepartmentID equals tbDrpartment.DepartmentID
                        join tbMessage in myModels.SYS_Message on tbDrpartment.MessageID equals tbMessage.MessageID

                        select new Staffdttcc
                        {
                            StaffID = tbStaff.StaffID,
                            DepartmentID = tbDrpartment.DepartmentID,
                            MessageID = tbMessage.MessageID,
                            StaffName = tbStaff.StaffName.Trim(),            //员工姓名
                            StaffNumber = tbStaff.StaffNumber.Trim(),        //员工编号
                            State = tbStaff.State.Trim(),                    //状态
                            Duty = tbStaff.Duty.Trim(),                      //职责
                            Sex = tbStaff.Sex.Trim(),                        //性别CultureLevel
                            ChineseName = tbMessage.ChineseName,      //所属组织
                            CultureLevel = tbStaff.CultureLevel.Trim(),       //文化程度
                            JoinJob = tbStaff.JoinJob.ToString(),      //入职日期
                            WhetherStart = tbStaff.WhetherStart,       //是否启用
                        }).ToList();
            #region MyRegion


            if (!string.IsNullOrEmpty(StaffName))
            {
                list = list.Where(m => m.StaffName.Contains(StaffName)).ToList();
            }
            if (!string.IsNullOrEmpty(StaffNumber))
            {
                list = list.Where(m => m.StaffNumber.Contains(StaffNumber)).ToList();
            }
            if (!string.IsNullOrEmpty(State))
            {
                list = list.Where(m => m.State == State).ToList();
            }
            if (!string.IsNullOrEmpty(Duty))
            {
                list = list.Where(m => m.StaffNumber.Contains(Duty)).ToList();
            }
            if (!string.IsNullOrEmpty(Sex))
            {
                list = list.Where(m => m.Sex == Sex).ToList();
            }
            if (!string.IsNullOrEmpty(ChineseName))
            {
                list = list.Where(m => m.StaffNumber.Contains(ChineseName)).ToList();
            }
            if (!string.IsNullOrEmpty(CultureLevel))
            {
                list = list.Where(m => m.StaffName.Contains(CultureLevel)).ToList();
            }
            if (!string.IsNullOrEmpty(JoinJob))
            {
                list = list.Where(m => m.StaffNumber.Contains(JoinJob)).ToList();
            }

            if (StaffID > 0)
            {
                list = list.Where(m => m.StaffID == StaffID).ToList();
            }
            if (A == 1)//分流，返回列表
            {
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            #endregion
            #endregion
            //将查询出来的结果转化为DataTable的格式
            DataTable dtResult = LINQToDataTable(list);
            //1、实例化数据集aa
            pp.ccc dbReport = new pp.ccc();
            //2、将dtResult放入数据集中名为"Message"的表格中  与数据集的表相同
            dbReport.Tables["Staff"].Merge(dtResult);
            //3、实例化数据报表
            pp.Staff rp = new pp.Staff();
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

        /// <summary>
        /// 导出
        /// </summary>
        /// <returns></returns>
        public ActionResult DeriveStaffControl(int? StaffID, string StaffName, string StaffNumber, string State, string Duty, string Sex, string ChineseName, string CultureLevel, string JoinJob, int? A)
        {
            #region MyRegion


            var list = (from tbStaff in myModels.SYS_Staff
                        join tbDrpartment in myModels.SYS_Department on tbStaff.DepartmentID equals tbDrpartment.DepartmentID
                        join tbMessage in myModels.SYS_Message on tbDrpartment.MessageID equals tbMessage.MessageID

                        select new Staffdttcc
                        {
                            StaffID = tbStaff.StaffID,
                            DepartmentID = tbDrpartment.DepartmentID,
                            MessageID = tbMessage.MessageID,
                            StaffName = tbStaff.StaffName.Trim(),                //员工姓名
                            StaffNumber = tbStaff.StaffNumber.Trim(),              //员工编号
                            State = tbStaff.State.Trim(),                    //状态
                            Duty = tbStaff.Duty.Trim(),                     //职责
                            Sex = tbStaff.Sex.Trim(),                      //性别CultureLevel
                            ChineseName = tbMessage.ChineseName,                   //所属组织
                            CultureLevel = tbStaff.CultureLevel.Trim(),             //文化程度
                            JoinJob = tbStaff.JoinJob.ToString(),              //入职日期
                            WhetherStart = tbStaff.WhetherStart,                    //是否启用
                        }).ToList();


            if (!string.IsNullOrEmpty(StaffName))
            {
                list = list.Where(m => m.StaffName.Contains(StaffName)).ToList();
            }
            if (!string.IsNullOrEmpty(StaffNumber))
            {
                list = list.Where(m => m.StaffNumber.Contains(StaffNumber)).ToList();
            }
            if (!string.IsNullOrEmpty(State))
            {
                list = list.Where(m => m.State == State).ToList();
            }
            if (!string.IsNullOrEmpty(Duty))
            {
                list = list.Where(m => m.StaffNumber.Contains(Duty)).ToList();
            }
            if (!string.IsNullOrEmpty(Sex))
            {
                list = list.Where(m => m.Sex == Sex).ToList();
            }
            if (!string.IsNullOrEmpty(ChineseName))
            {
                list = list.Where(m => m.StaffNumber.Contains(ChineseName)).ToList();
            }
            if (!string.IsNullOrEmpty(CultureLevel))
            {
                list = list.Where(m => m.StaffName.Contains(CultureLevel)).ToList();
            }
            if (!string.IsNullOrEmpty(JoinJob))
            {
                list = list.Where(m => m.StaffNumber.Contains(JoinJob)).ToList();
            }

            if (StaffID > 0)
            {
                list = list.Where(m => m.StaffID == StaffID).ToList();
            }
            if (A == 1)//分流，返回列表
            {
                return Json(list, JsonRequestBehavior.AllowGet);
            }
            #endregion

            //1、创建Excel工作簿
            NPOI.HSSF.UserModel.HSSFWorkbook Leave = new NPOI.HSSF.UserModel.HSSFWorkbook();
            //2、创建工作表
            // NPOI.SS.UserModel.ISheet sheet = Leave.CreateSheet("");
            NPOI.SS.UserModel.ISheet sheet = Leave.CreateSheet("南派三叔");
            //3、设计表头
            //3.1 创建第一行
            NPOI.SS.UserModel.IRow row = sheet.CreateRow(0);
            //3.2 设计字段
            row.CreateCell(0).SetCellValue("员工姓名");
            row.CreateCell(1).SetCellValue("员工编号");
            row.CreateCell(2).SetCellValue("状态");
            row.CreateCell(3).SetCellValue("职责");
            row.CreateCell(4).SetCellValue("性别");
            row.CreateCell(5).SetCellValue("所属组织");
            row.CreateCell(6).SetCellValue("文化程度");
            row.CreateCell(7).SetCellValue("入职日期");
            row.CreateCell(8).SetCellValue("是否启用");

            //4、循环写入数据
            for (var i = 0; i < list.Count(); i++)
            {
                NPOI.SS.UserModel.IRow lll = sheet.CreateRow(i + 1);
                lll.CreateCell(0).SetCellValue(list[i].StaffName);
                lll.CreateCell(1).SetCellValue(list[i].StaffNumber);
                lll.CreateCell(2).SetCellValue(list[i].State);
                lll.CreateCell(3).SetCellValue(list[i].Duty);
                lll.CreateCell(4).SetCellValue(list[i].Sex);
                lll.CreateCell(5).SetCellValue(list[i].ChineseName);
                lll.CreateCell(6).SetCellValue(list[i].CultureLevel);
                lll.CreateCell(7).SetCellValue(list[i].JoinJob);
                lll.CreateCell(8).SetCellValue(list[i].WhetherStart);
            }
            //5、将Excel表格转化为文件流输出
            MemoryStream DulaDula = new MemoryStream();
            Leave.Write(DulaDula);
            //输出之前调用Seek（偏移量，游标位置）方法：确定流开始的位置
            DulaDula.Seek(0, SeekOrigin.Begin);
            //6、文件名
            string fileName = "南派三叔" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            return File(DulaDula, "application/vnd.ms-excel", fileName);

        }

        #endregion
    }
}

using Hotel.VO;
using NPOI.HSSF.UserModel;
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

    public class BranchManagementController : Controller
    {
        // GET: PPAP/BranchManagement
        Models.SeaTransportationEntities myModels = new SeaTransportationEntities();
        public ActionResult BranchManagement()
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

        #region 部门管理
        /// <summary>
        /// 树形
        /// </summary>
        /// <returns></returns>
        public ActionResult tree()
        {
            var list = (from tbDrpartment in myModels.SYS_Department
                        select new
                        {
                            id = tbDrpartment.DepartmentID,
                            name = tbDrpartment.DepartmentName
                        }).Distinct().ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 查询
        /// </summary>
        /// <param name="bsgridPage"></param>
        /// <returns></returns>
        public ActionResult SelectDepartment(BsgridPage bsgridPage, short? MessageID, string DepartmentName, string DepartmentMobile, string InteriorPhone, string DepartmentTop, string DepartmentDuty, string DepartmentDescribe, string ChineseName,int? DepartmentID)
        {
            var Department = (from tbDepartment in myModels.SYS_Department
                              join tbMessage in myModels.SYS_Message on tbDepartment.MessageID equals tbMessage.MessageID

                              select new Departmen
                              {
                                  DepartmentID = tbDepartment.DepartmentID, //部门ID
                                  DepartmentName = tbDepartment.DepartmentName,//部门名称
                                  DepartmentMobile = tbDepartment.DepartmentMobile, //  联系电话
                                  InteriorPhone = tbDepartment.InteriorPhone,//内部电话 
                                  DepartmentTop = tbDepartment.DepartmentTop,//部门领导 
                                  DepartmentDuty = tbDepartment.DepartmentDuty,//部门职责
                                  DepartmentDescribe = tbDepartment.DepartmentDescribe,//部门描述
                                  ChineseName = tbMessage.ChineseName,//所属组织
                                  MessageID = tbDepartment.MessageID,//id 公司
                              }).ToList();

            if (MessageID > 0)
            {
                Department = Department.Where(m => m.MessageID == MessageID).ToList();
            }
            if (!string.IsNullOrEmpty(DepartmentName))
            {
                Department = Department.Where(m => m.DepartmentName.Contains(DepartmentName)).ToList();
                // ccc = ccc.Where(m => m.Describe.Contains(Describe)).ToList();//描述
            }
            if (!string.IsNullOrEmpty(DepartmentMobile))
            {
                Department = Department.Where(m => m.DepartmentMobile.Contains(DepartmentMobile)).ToList();
            }
            if (!string.IsNullOrEmpty(InteriorPhone))
            {
                Department = Department.Where(m => m.InteriorPhone.Contains(InteriorPhone)).ToList();
            }
            if (!string.IsNullOrEmpty(DepartmentTop))
            {
                Department = Department.Where(m => m.DepartmentTop.Contains(DepartmentTop)).ToList();
            }
            if (!string.IsNullOrEmpty(DepartmentDuty))
            {
                Department = Department.Where(m => m.DepartmentDuty.Contains(DepartmentDuty)).ToList();
            }
            if (!string.IsNullOrEmpty(DepartmentDescribe))
            {
                Department = Department.Where(m => m.DepartmentDuty.Contains(DepartmentDuty)).ToList();
            }
            if (!string.IsNullOrEmpty(ChineseName))
            {
                Department = Department.Where(m => m.ChineseName.Contains(ChineseName)).ToList();
            }
            if (DepartmentID > 0)
            {
                Department = Department.Where(m => m.DepartmentID == DepartmentID).ToList();
            }

            int pageCount = Department.Count();
            List<Departmen> wcg = Department.Skip(bsgridPage.GetStartIndex()).Take(bsgridPage.pageSize).ToList();
            Bsgrid<Departmen> bsgrid = new Bsgrid<Departmen>()
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
        public ActionResult InsertDepartment(short MessageID, string DepartmentName, string DepartmentMobile, string InteriorPhone, string DepartmentTop, string DepartmentDuty, string DepartmentDescribe, string ChineseName)
        {
            string strMsg = "failed";
            try
            {
                var listDepartment = (from tbDepartment in myModels.SYS_Department
                                      where tbDepartment.DepartmentMobile == DepartmentMobile || tbDepartment.DepartmentName == DepartmentName
                                      where tbDepartment.MessageID == MessageID
                                      select tbDepartment).Count();
                
                if (listDepartment == 0)
                {
                    SYS_Department myDepartment = new SYS_Department();
                    myDepartment.DepartmentName = DepartmentName.Trim();//部门名称
                    myDepartment.DepartmentMobile = DepartmentMobile.Trim();//  联系电话
                    myDepartment.InteriorPhone = InteriorPhone.Trim();//内部电话 
                    myDepartment.DepartmentTop = DepartmentTop.Trim();//部门领导 
                    myDepartment.DepartmentDuty = DepartmentDuty.Trim();//部门职责
                    myDepartment.DepartmentDescribe = DepartmentDescribe.Trim();//部门描述
                    myDepartment.MessageID = MessageID;
                    //将实例化的实体插入数据库实体ss

                    myModels.SYS_Department.Add(myDepartment);
                    //保存发生变化后的数据库
                    myModels.SaveChanges();

                    strMsg = "新增成功！";
                }
                else
                {
                    strMsg = "该部门信息已经存在，不需要重复输入数据！";
                    return Json("success", JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception e)
            {
                strMsg = "failed";
            }
            return Json(strMsg, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 修改
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateDepartment(short MessageID, int DepartmentID, string DepartmentName, string DepartmentMobile, string InteriorPhone, string DepartmentTop, string DepartmentDuty, string DepartmentDescribe, string ChineseName)
        {
            string strMsg = "failed";
            try
            {
                Models.SYS_Department myDepartment = myModels.SYS_Department.Where(p => p.DepartmentID == DepartmentID).Single();
                myDepartment.DepartmentName = DepartmentName.Trim();//部门名称
                myDepartment.DepartmentMobile = DepartmentMobile.Trim();//  联系电话
                myDepartment.InteriorPhone = InteriorPhone.Trim();//内部电话 
                myDepartment.DepartmentTop = DepartmentTop.Trim();//部门领导 
                myDepartment.DepartmentDuty = DepartmentDuty.Trim();//部门职责
                myDepartment.DepartmentDescribe = DepartmentDescribe.Trim();//部门描述
                myDepartment.MessageID = MessageID;
                //获取和实质对象实体的状态 = EntityState的枚举值
                myModels.Entry(myDepartment).State = System.Data.Entity.EntityState.Modified;
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
        /// <param name="DepartmentID"></param>
        /// <returns></returns>
        public ActionResult DelectDepartment(int? DepartmentID)
        {
            string strMsg = "fail";

            try
            {
                var Departement = myModels.SYS_Department.Where(m => m.DepartmentID == DepartmentID).Single();

                myModels.SYS_Department.Remove(Departement);
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

        public ActionResult DeriveDepartment(short? MessageID, string DepartmentName, string DepartmentMobile, string InteriorPhone, string DepartmentTop, string DepartmentDuty, string DepartmentDescribe, string ChineseName)
        {

            var Department = (from tbDepartment in myModels.SYS_Department
                              join tbMessage in myModels.SYS_Message on tbDepartment.MessageID equals tbMessage.MessageID

                              select new Departmen
                              {
                                  DepartmentID = tbDepartment.DepartmentID, //部门ID
                                  DepartmentName = tbDepartment.DepartmentName,//部门名称
                                  DepartmentMobile = tbDepartment.DepartmentMobile, //  联系电话
                                  InteriorPhone = tbDepartment.InteriorPhone,//内部电话 
                                  DepartmentTop = tbDepartment.DepartmentTop,//部门领导 
                                  DepartmentDuty = tbDepartment.DepartmentDuty,//部门职责
                                  DepartmentDescribe = tbDepartment.DepartmentDescribe,//部门描述
                                  ChineseName = tbMessage.ChineseName,//所属组织
                                  MessageID = tbDepartment.MessageID,//id 公司
                              }).ToList();

            if (MessageID > 0)
            {
                Department = Department.Where(m => m.MessageID == MessageID).ToList();
            }
            if (!string.IsNullOrEmpty(DepartmentName))
            {
                Department = Department.Where(m => m.DepartmentName.Contains(DepartmentName)).ToList();
                // ccc = ccc.Where(m => m.Describe.Contains(Describe)).ToList();//描述
            }
            if (!string.IsNullOrEmpty(DepartmentMobile))
            {
                Department = Department.Where(m => m.DepartmentMobile.Contains(DepartmentMobile)).ToList();
            }
            if (!string.IsNullOrEmpty(InteriorPhone))
            {
                Department = Department.Where(m => m.InteriorPhone.Contains(InteriorPhone)).ToList();
            }
            if (!string.IsNullOrEmpty(DepartmentTop))
            {
                Department = Department.Where(m => m.DepartmentTop.Contains(DepartmentTop)).ToList();
            }
            if (!string.IsNullOrEmpty(DepartmentDuty))
            {
                Department = Department.Where(m => m.DepartmentDuty.Contains(DepartmentDuty)).ToList();
            }
            if (!string.IsNullOrEmpty(DepartmentDescribe))
            {
                Department = Department.Where(m => m.DepartmentDuty.Contains(DepartmentDuty)).ToList();
            }
            if (!string.IsNullOrEmpty(ChineseName))
            {
                Department = Department.Where(m => m.ChineseName.Contains(ChineseName)).ToList();
            }

            //将查询出来的数据转化为对象列表的格式
            List<Departmen> listEEe = Department.ToList();

            //创建Excel工作簿
            HSSFWorkbook excelBook = new HSSFWorkbook();
            //在工作簿中创建工作表并为工作表命名
            NPOI.SS.UserModel.ISheet sheet = excelBook.CreateSheet("考生信息");
            //创建表头
            NPOI.SS.UserModel.IRow rowl = sheet.CreateRow(0);
            //写入表格的字段
            rowl.CreateCell(0).SetCellValue("部门名称");
            rowl.CreateCell(1).SetCellValue("联系电话");
            rowl.CreateCell(2).SetCellValue("内部电话");
            rowl.CreateCell(3).SetCellValue("部门领导");
            rowl.CreateCell(4).SetCellValue("部门职责");
            rowl.CreateCell(5).SetCellValue("部门描述");
            rowl.CreateCell(6).SetCellValue("部门编号");
            rowl.CreateCell(7).SetCellValue("所属组织");
            for (var i = 0; i < listEEe.Count; i++)
            {
                //创建行
                NPOI.SS.UserModel.IRow rowTemp = sheet.CreateRow(i + 1);
                //写入数据
                rowTemp.CreateCell(0).SetCellValue(listEEe[i].DepartmentName);
                rowTemp.CreateCell(1).SetCellValue(listEEe[i].DepartmentMobile);
                rowTemp.CreateCell(2).SetCellValue(listEEe[i].InteriorPhone);
                rowTemp.CreateCell(3).SetCellValue(listEEe[i].DepartmentTop);
                rowTemp.CreateCell(4).SetCellValue(listEEe[i].DepartmentDuty);
                rowTemp.CreateCell(5).SetCellValue(listEEe[i].DepartmentDescribe);
                rowTemp.CreateCell(6).SetCellValue(listEEe[i].ChineseName);
            }
            //文件名
            var fileName = "太宰治" + DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss-ffff") + ".xls";
            //将Excel表格转化为流，输出
            //创建文件流
            System.IO.MemoryStream bookStream = new System.IO.MemoryStream();
            //文件写入流（向流中写入字节序列）
            excelBook.Write(bookStream);
            //输出之前调用Seek（偏移量，游标位置) 把0位置指定为开始位置
            bookStream.Seek(0, System.IO.SeekOrigin.Begin);
            //MIME文件类型(Multipurpose Internet Mail Extensions)多用途互联网邮件扩展类型
            return File(bookStream, "application/vnd.ms-excel", fileName);
        }

        /// <summary>
        /// 打印
        /// </summary>
        /// <returns></returns>
        public ActionResult StampDepartment(short? MessageID, string DepartmentName, string DepartmentMobile, string InteriorPhone, string DepartmentTop, string DepartmentDuty, string DepartmentDescribe, string ChineseName)
        {
            #region MyRegion


            var Department = (from tbDepartment in myModels.SYS_Department
                              join tbMessage in myModels.SYS_Message on tbDepartment.MessageID equals tbMessage.MessageID

                              select new Departmen
                              {
                                  DepartmentID = tbDepartment.DepartmentID, //部门ID
                                  DepartmentName = tbDepartment.DepartmentName,//部门名称
                                  DepartmentMobile = tbDepartment.DepartmentMobile, //  联系电话
                                  InteriorPhone = tbDepartment.InteriorPhone,//内部电话 
                                  DepartmentTop = tbDepartment.DepartmentTop,//部门领导 
                                  DepartmentDuty = tbDepartment.DepartmentDuty,//部门职责
                                  DepartmentDescribe = tbDepartment.DepartmentDescribe,//部门描述
                                  ChineseName = tbMessage.ChineseName,//所属组织
                                  MessageID = tbDepartment.MessageID,//id 公司
                              }).ToList();

            if (MessageID > 0)
            {
                Department = Department.Where(m => m.MessageID == MessageID).ToList();
            }
            if (!string.IsNullOrEmpty(DepartmentName))
            {
                Department = Department.Where(m => m.DepartmentName.Contains(DepartmentName)).ToList();
                // ccc = ccc.Where(m => m.Describe.Contains(Describe)).ToList();//描述
            }
            if (!string.IsNullOrEmpty(DepartmentMobile))
            {
                Department = Department.Where(m => m.DepartmentMobile.Contains(DepartmentMobile)).ToList();
            }
            if (!string.IsNullOrEmpty(InteriorPhone))
            {
                Department = Department.Where(m => m.InteriorPhone.Contains(InteriorPhone)).ToList();
            }
            if (!string.IsNullOrEmpty(DepartmentTop))
            {
                Department = Department.Where(m => m.DepartmentTop.Contains(DepartmentTop)).ToList();
            }
            if (!string.IsNullOrEmpty(DepartmentDuty))
            {
                Department = Department.Where(m => m.DepartmentDuty.Contains(DepartmentDuty)).ToList();
            }
            if (!string.IsNullOrEmpty(DepartmentDescribe))
            {
                Department = Department.Where(m => m.DepartmentDuty.Contains(DepartmentDuty)).ToList();
            }
            if (!string.IsNullOrEmpty(ChineseName))
            {
                Department = Department.Where(m => m.ChineseName.Contains(ChineseName)).ToList();
            }
            #endregion
            //将查询出来的结果转化为DataTable的格式
            DataTable dtResult = LINQToDataTable(Department);
            //1、实例化数据集
            pp.ccc dbReport = new pp.ccc();
            //2、将dtResult放入数据集中名为"Message"的表格中  与数据集的表相同
            dbReport.Tables["Department"].Merge(dtResult);
            //3、实例化数据报表
            pp.Department rp = new pp.Department();
            //4、获取报表的物理文件路径
            string strRpPath = System.Web.HttpContext.Current.Server.MapPath("~/") + "Areas\\PPAP\\pp\\Department.rpt";
            //5、将报表加载到报表模板中
            rp.Load(strRpPath);
            //6、设置报表的数据源
            rp.SetDataSource(dbReport);
            //7、把报表转化为文件流输出
            Stream dbStream = rp.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
            return File(dbStream, "application/pdf");
        }

        /// <summary>
        /// 美剧
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="varlist"></param>
        /// <returns></returns>
        public DataTable LINQToDataTable<T>(IEnumerable<T> varlist)
        {
            //定义要返回的DataTable对象
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

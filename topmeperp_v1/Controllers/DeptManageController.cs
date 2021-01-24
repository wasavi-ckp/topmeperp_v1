using log4net;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;

namespace topmeperp.Controllers
{
    public class DeptManageController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        DepartmentManage service = new DepartmentManage();
        // GET: UserManage
        public ActionResult Index()
        {
            logger.Info("index");
            //1.是公司名稱
            SelectList dept = new SelectList(service.getDepartmentOrg(1), "DEP_ID", "DEPT_NAME");
            ViewBag.d_parentId = dept;
            ViewBag.TreeString = service.getDepartment4Tree(0);
            return View();
        }
        [HttpPost]
        public ActionResult Query(FormCollection form)
        {
            logger.Info("criteria dpep_name=" + form.Get("userid"));
            return PartialView("DepartmentList", service.Org_department);
            //return View(userService.userManageModels);
        }
        //新增或修改部門
        public String addDepartment(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            //懶得把Form綁SYS_USER 直接先把Form 值填滿
            ENT_DEPARTMENT d = new ENT_DEPARTMENT();
            if (form.Get("d_depId").Trim() != "")
            {
                d.DEP_ID = Convert.ToInt64(form.Get("d_depId").Trim());
            }
            d.DEPT_CODE = form.Get("d_deptCode").Trim();
            d.DEPT_NAME = form.Get("d_deptName").Trim();
            d.MANAGER = form.Get("d_Manager").Trim();
            d.DESC = form.Get("d_desc").Trim();
            if (null != form.Get("d_parentId") && "" != form.Get("d_parentId").Trim())
            {
                d.PARENT_ID = Convert.ToInt64(form.Get("d_parentId").Trim());
            }
            SYS_USER loginUser = (SYS_USER)Session["user"];
            d.CREATE_ID = loginUser.USER_ID;
            d.CREATE_DATE = DateTime.Now;
            msg = "新增部門成功(" + service.addDepartment(d) + ")";

            //logger.Info("Request:user_ID=" + form["u_userid"]);
            return msg;
        }

        //刪除部門
        public String delDepartment(FormCollection form)
        {
            long depId = Convert.ToInt64(form.Get("d_depId").Trim());
            SYS_USER loginUser = (SYS_USER)Session["user"];
            logger.Info("User " + loginUser.USER_ID + " delete departnt :" + depId + ",Name" + form.Get("d_deptName"));
            string msg = "";
            try
            {
                service.delDepartment(depId);
                msg = "刪除部門成功";
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message + ":" + ex.StackTrace);
                msg = "刪除部門失敗";
            }
            return msg;
        }

        //取得某一Dept 基本資料
        public string getDept()
        {
            string deptString = Request["deptid"];
            logger.Info("get Dept id=" + deptString);
            string[] deptId = deptString.Split('-');

            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string deptJson = objSerializer.Serialize(service.getDepartmentById(deptId[0]));
            logger.Info("dept info=" + deptJson);
            return deptJson;
        }
    }
}

using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;

namespace topmeperp.Controllers
{
    public class UserManageController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(UserManageController));
        UserManage userService = new UserManage();
        // GET: UserManage
        public ActionResult Index()
        {
            log.Info("index");
            userService.getAllRole();
            ViewData.Add("roles", userService.userManageModels.sysRole);
            SelectList roles = new SelectList(userService.userManageModels.sysRole, "ROLE_ID", "ROLE_NAME");
            ViewBag.roles = roles;
            //將資料存入TempData 減少不斷讀取資料庫
            TempData.Remove("roles");
            TempData.Add("roles", userService.userManageModels.sysRole);

            //DepartmentManage
            DepartmentManage ds = new DepartmentManage();
            SelectList dep_code = new SelectList(ds.getDepartmentOrg(1), "DEPT_CODE", "DEPT_NAME");
            ViewBag.dep_code = dep_code;
            return View();
        }
        [HttpPost]
        public ActionResult Query(FormCollection form)
        {
            log.Info("criteria user_id=" + form.Get("userid") + ",username=" + form.Get("username") + ",tel=" + form.Get("tel") + ",roleid=" + form.Get("roles"));
            SYS_USER u_user = new SYS_USER();
            u_user.USER_ID = form.Get("userid");
            u_user.USER_NAME = form.Get("username");
            u_user.TEL = form.Get("tel");
            userService.getUserByCriteria(u_user, form.Get("roles").Trim());
            ViewBag.SearchResult = "User 共" + userService.userManageModels.sysUsers.Count() + "筆資料!!";
            //回傳部分網頁
            return PartialView("UserList", userService.userManageModels);
            //return View(userService.userManageModels);
        }
        //新增或修改使用者
        public String addUser(FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            //懶得把Form綁SYS_USER 直接先把Form 值填滿
            SYS_USER u = new SYS_USER();
            u.USER_ID = form.Get("u_userid").Trim();
            u.USER_NAME = form.Get("u_name").Trim();
            u.PASSWORD = form.Get("u_password").Trim();
            u.EMAIL = form.Get("u_email").Trim();
            u.TEL = form.Get("u_tel").Trim();
            u.TEL_EXT = form.Get("u_tel_ext").Trim();
            u.MOBILE = form.Get("u_mobile").Trim();
            u.ROLE_ID = form.Get("roles").Trim();
            u.DEP_CODE = form.Get("dep_code").Trim();
            SYS_USER loginUser = (SYS_USER)Session["user"];
            u.CREATE_ID = loginUser.USER_ID;
            u.CREATE_DATE = DateTime.Now;
            int i = userService.addNewUser(u);
            if (i == 0)
            {
                msg = userService.message;
            }
            else
            {
                msg = "帳號更新成功";
            }

            log.Info("Request:user_ID=" + form["u_userid"]);
            return msg;
        }
        //取得某一User 基本資料
        public string getUser(string userid)
        {
            log.Info("get user id=" + userid);
            SYS_USER u = userService.getUser(userid);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string userJson = objSerializer.Serialize(u);
            log.Info("user info=" + userJson);
            return userJson;
        }
        public ActionResult userProfile()
        {
            SYS_USER u = (SYS_USER)Session["user"];
            return View("UserProfile", u);
        }
        [HttpPost]
        public ActionResult userProfile(SYS_USER newUser)
        {
            SYS_USER u = (SYS_USER)Session["user"];
            if (null!= newUser.PASSWORD && newUser.PASSWORD != "")
            {
                //check password
                if (null != newUser.PASSWORD && newUser.PASSWORD != Request["confirmpwd"])
                {
                    ViewBag.ErrorMessage = "密碼欄位有問題，請重新輸入!";
                    return View("UserProfile", u);
                }
                u.PASSWORD = newUser.PASSWORD;
            }
            u.USER_NAME = newUser.USER_NAME;
            u.TEL= newUser.TEL;
            u.TEL_EXT = newUser.TEL_EXT;
            u.EMAIL = newUser.EMAIL;
            u.FAX = newUser.FAX;
            u.MOBILE = newUser.MOBILE;

            u.MODIFY_ID = u.USER_ID;
            u.MODIFY_DATE = DateTime.Now;
            int i = userService.updateProfile(u);
            Session["user"] = u;
            ViewBag.Message = "更改個人資料成功!!";
            return View("UserProfile", u);
        }

    }
}
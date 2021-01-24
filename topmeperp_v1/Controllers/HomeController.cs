using log4net;
using System;
using System.Collections.Generic;
using System.Web.Mvc;
using topmeperp.Filter;
using topmeperp.Models;
using topmeperp.Service;

namespace topmeperp.Controllers
{
    [AuthFilter]
    public class HomeController : Controller
    {
        private static log4net.ILog Log { get; set; }
        ILog log = log4net.LogManager.GetLogger(typeof(HomeController));

        [AllowAnonymous]
        public ActionResult Index()
        {
            log.Info("log4net test!!");
            return View();
        }

        //
        // GET: /Home/Login
        [AllowAnonymous]
        //List<SYS_FUNCTION> functions = null;

        public ActionResult Login(string returnUrl)
        {
            log.Info("log4net test Login by get!!");
            ViewBag.ReturnUrl = returnUrl;
            return View();
        }
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        UserService u = null;
        //
        // POST: /Home/Login
        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public ActionResult Login(SYS_USER model, string returnUrl)
        {
            log.Info("log4net test Login by post:" + model.USER_ID + "," + model.PASSWORD);
            getPrivilegeByUser(model.USER_ID, model.PASSWORD);
            //2.檢查權限是否存在
            if (null == u.loginUser)
            {
                //2.1 當帳號不存在時，將View 設回首頁同時帶入錯誤訊息               
                log.Info("Login fail");
                ViewBag.ErrorMessage = "帳號密碼有誤，請洽系統管理者!!";
                return View();

            }
            else
            {
                //3.登入成功導入功能主畫面
                log.Info("Login Success by :" + model.USER_ID);
                List<SYS_FUNCTION> lst = (List<SYS_FUNCTION>)Session["functions"];
                //4.假如沒有授權資料，導回登入頁說明錯誤訊息
                if (null == lst || lst.Count == 0)
                {
                    log.Info("not privilege");
                    ViewBag.ErrorMessage = "您的權限有問題，請洽系統管理者!!";
                    return View();
                }
                else
                {
                    log.Debug("redirec to first function=" + lst[0].FUNCTION_URI);
                    return Redirect(lst[0].FUNCTION_URI);
                }
            }

        }

        private void getPrivilegeByUser(String userid, String passwd)
        {
            u = new UserService();
            u.Login(userid, passwd);
            Session.Add("user", u.loginUser);
            Session.Add("functions", u.userPrivilege);
        }
        public ActionResult Logout()
        {
            SYS_USER u = (SYS_USER)Session["user"];
            log.Info(u.USER_ID + " Logout!!");
            //1.清空Session 
            Session.RemoveAll();
            //2.導回登入頁
            return RedirectToAction("Login", "Home");

        }
    }
}
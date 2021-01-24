using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;

namespace topmeperp.Controllers
{
    public class RoleManageController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(UserManageController));
        UserManage userService = new UserManage();
        // GET: RoleManage
        public ActionResult Index()
        {
            log.Info("index");
            userService.getAllRole();
            SelectList roles = new SelectList(userService.userManageModels.sysRole, "ROLE_ID", "ROLE_NAME");
            ViewBag.roles = roles;
            return View(userService.getPrivilege(""));
        }
        public ActionResult FunctionList()
        {
            string roleid = Request["roles"];
            log.Debug(Request.IsAjaxRequest());
            log.Info("index roleid=" + roleid);
            return PartialView(userService.getPrivilege(roleid));
        }
        public string UpdatePrivilege()
        {
            log.Info("new privilege:" + Request["hadPrivilege"]);
            string roleid = Request.Params["roleid"];
            string[] funs=new string[0];
            if (null != Request.Params["hadPrivilege"])
            {
                funs = Request.Params["hadPrivilege"].Split(',');
            }
            userService.updatePrivilege(roleid, funs);
            return "權限已經更新!!";
        }
        public string addOrUpdateRole()
        {
            log.Info("RoleInfo:" + Request["r_roleid"] + ",Name=" + Request["r_name"]);
            SYS_ROLE r = new SYS_ROLE();
            r.ROLE_ID = Request["r_roleid"];
            r.ROLE_NAME = Request["r_name"];

            if (userService.addOrUpdateRole(r) > 0)
            {
                return "更新成功!!";
            }
            else
            {
                return "更新失敗!!";
            }

        }
    }
}
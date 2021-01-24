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
    /// <summary>
    /// 專案組織管理..用於處理專案相關聯繫人
    /// </summary>
    public class ProjectOrganizationController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        TnderProjectService service = new TnderProjectService();
        // GET: ProjectOrganization
        public ActionResult Index(string id)
        {
            string projectId = id;
            logger.Info("project detail page projectid = " + projectId);
            ViewBag.projectid = projectId;
            TndProjectModels viewModel = new TndProjectModels();
            List<TND_TASKASSIGN> lstTask = null;

            lstTask = service.getTaskByPrjId(projectId,null);
            TND_PROJECT p = service.getProjectById(projectId);
            //?
            ViewBag.taskAssign = projectId;
            viewModel.tndProject = p;
            viewModel.tndTaskAssign = lstTask;

            //畫面上權限管理控制
            //頁面上使用ViewBag 定義開關\@ViewBag.F00003
            //由Session 取得權限清單
            List<SYS_FUNCTION> lstFunctions = (List<SYS_FUNCTION>)Session["functions"];
            //開關預設關閉
            @ViewBag.F00003 = "disabled";
            //輪巡功能清單，若全線存在則將開關打開 @ViewBag.F00003 = "";
            foreach (SYS_FUNCTION f in lstFunctions)
            {
                if (f.FUNCTION_ID == "F00003")
                {
                    @ViewBag.F00003 = "";
                }
            }
            return View(viewModel);
        }
        public string delTaskItem()
        {
            TND_TASKASSIGN t = new TND_TASKASSIGN();
            t.TASK_ID = long.Parse(Request["itemid"].ToString());
            SYS_USER u = (SYS_USER)Session["user"];
            service.delTask(u,t);
            return service.message;
        }
    }
}
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
    public class FormManageController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        //廠商計價單管理
        public ActionResult Index()
        {
            SelectList lstProject = new SelectList(PlanService.SearchProjectByName("", "專案執行",null), "PROJECT_ID", "PROJECT_NAME");
            //取得表單狀態參考資料
            SelectList status = new SelectList(SystemParameter.getSystemPara("ExpenseForm"), "KEY_FIELD", "VALUE_FIELD");
            ViewData.Add("status", status);
            ViewData.Add("projects", lstProject);
            return View();
        }
        //查詢廠商計價單
        public ActionResult FormList()
        {
            logger.Info("Search For Estimation Form !!");
            TnderProjectService tndservice = new TnderProjectService();
            string projectid = Request["projects"];
            Flow4Estimation s = new Flow4Estimation();
            List<ExpenseFlowTask> lstEST = s.getEstimationFormRequest(null, null, null, projectid, Request["status"]);
            return PartialView("_EstimationFormList", lstEST);
        }
    }
}
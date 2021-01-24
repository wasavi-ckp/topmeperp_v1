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
    public class CostChangeController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        // GET: //成本異動採購作業(針對已經完成審核之異動單進行採購作業
        public ActionResult Index()
        {
            SYS_USER u = (SYS_USER)Session["user"];
            CostChangeService cs = new CostChangeService();
            SelectList lstProject = new SelectList(PlanService.SearchProjectByName("", "專案執行",u), "PROJECT_ID", "PROJECT_NAME");
            ViewData.Add("projects", lstProject);
            return View();
        }
        //查詢異動單
        public ActionResult FormList()
        {
            string projectId = Request["projects"];
            string remark = Request["remark"];
            string noInquiry = Request["noInquiry"];
            CostChangeService cs = new CostChangeService();
            //取得通過審核之異動單資料 STATUS=30
            List<CostChangeForm> lst = cs.getCostChangeForm(projectId, "30", remark, noInquiry);
            ViewBag.SearchResult = "共取得" + lst.Count + "筆資料!!";
            return PartialView("_ChangeFormList",lst);
        }
        //成本異動單表單
        public ActionResult costChangeForm(string id)
        {
            string formId = id;
            logger.Debug("formId=" + formId);
            //先取得資料
            CostChangeService cs = new CostChangeService();
            cs.getChangeOrderForm(formId);

            ViewBag.FormId = formId;
            ViewBag.projectId = cs.project.PROJECT_ID;
            ViewBag.projectName = cs.project.PROJECT_NAME;
            ViewBag.settlementDate = cs.form.SETTLEMENT_DATE;
            //取得表單資料存入Session
            Flow4CostChange wfs = new Flow4CostChange();
            wfs.getTask(formId);
            wfs.task.FormData = cs.form;
            wfs.task.lstItem = cs.lstItem;
            Session["process"] = wfs.task;
            SelectList reasoncode = new SelectList(SystemParameter.getSystemPara("COSTHANGE", "REASON"), "KEY_FIELD", "VALUE_FIELD", cs.form.REASON_CODE);
            ViewData.Add("reasoncode", reasoncode);
            //財務處理區塊
            SelectList methodcode = new SelectList(SystemParameter.getSystemPara("COSTHANGE", "METHOD"), "KEY_FIELD", "VALUE_FIELD", cs.form.METHOD_CODE);
            ViewData.Add("methodcode", methodcode);
            return View(wfs.task);
        }
        /// <summary>
        /// 建立詢價單並轉至詢價單頁面
        /// </summary>
        public void createInquiryOrder()
        {
            string formId = Request["formId"];
            logger.Debug("formId=" + formId);
            SYS_USER u = (SYS_USER)Session["user"];
            CostChangeService cs = new CostChangeService();
            string inquiryFormId= cs.createInquiryOrderByChangeForm(formId,u);
            string url = "~/PurchaseForm/SinglePrjForm/" + inquiryFormId + "?update=Y";
            logger.Debug("Redirector:" + url);
            Response.Redirect(url);
        }
    }
}
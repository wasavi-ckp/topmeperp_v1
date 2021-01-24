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
    public class PaymentManagementController : EstimationController
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        // Post: PaymentManagement
        [HttpGet]
        public ActionResult Index(string id)
        {
            logger.Info("Search For Estimation Form !!");
            ViewBag.projectid = id;
            getProject(id);
            //取得表單狀態參考資料
            SelectList LstStatus = new SelectList(SystemParameter.getSystemPara("ExpenseForm"), "KEY_FIELD", "VALUE_FIELD");
            ViewData.Add("status", LstStatus);
            string strStatus = null;
            if (null != Request["status"])
            {
                strStatus = Request["status"];
            }else
            {
                strStatus = "30";
            }
            EstimationFormApprove model = new EstimationFormApprove();
            //取得審核中表單資料
            Flow4Estimation workflow = new Flow4Estimation();
            EstimationService es = new EstimationService();
            model.lstEstimationFlowTask = workflow.getEstimationFormRequest(Request["contractid"], Request["payee"], Request["estid"], id, strStatus);
            ViewBag.SearchResult = "共取得" + model.lstEstimationFlowTask + "筆資料";
            getProject(id);
            return View(model);
        }
        public void downLoadExpenseForm()
        {
            string formid = Request["formid"];
            logger.Debug("download Excel:"+ formid);
            EstimationService service = new EstimationService();
            ContractModels constract = service.getEstimationOrder(formid);
            if (null != constract)
            {
                PaymentExpenseFormToExcel poi = new PaymentExpenseFormToExcel();
                //檔案位置
                string fileLocation = poi.exportExcel(constract);
                //檔案名稱 HttpUtility.UrlEncode預設會以UTF8的編碼系統進行QP(Quoted-Printable)編碼，可以直接顯示的7 Bit字元(ASCII)就不用特別轉換。
                string filename = HttpUtility.UrlEncode(Path.GetFileName(fileLocation));
                Response.Clear();
                Response.Charset = "utf-8";
                Response.ContentType = "text/xls";
                Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", filename));
                ///"\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\" + form.FORM_ID + ".xlsx"
                Response.WriteFile(fileLocation);
                Response.End();
            }
        }
    }
}
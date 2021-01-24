using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;
using System.IO;
using Newtonsoft.Json;
using System.Data;

namespace topmeperp.Controllers
{
    public class EstimationController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        Service4Budget service = new Service4Budget();

        // GET: Estimation
        [topmeperp.Filter.AuthFilter]
        public ActionResult Index()
        {
            SYS_USER u = (SYS_USER)Session["user"];
            List<ProjectList> lstProject = PlanService.SearchProjectByName("", "專案執行','保固", u);
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            return View(lstProject);
        }

        //廠商估驗計價
        public ActionResult Valuation(string id)
        {
            logger.Info("valuation index : projectid=" + id);
            ViewBag.projectId = id;
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            return View();
        }
        public ActionResult queryInfo4Estimation()
        {
            ViewBag.projectId = Request["projectId"];
            ViewBag.projectName = Request["projectName"];

            EstimationService service = new EstimationService();
            ContractModels contract = new ContractModels();
            List<plansummary> lstContract = null;
            lstContract = service.getAllPlanContract(Request["projectId"]);
            contract.contractItems = lstContract;
            //由日報彙整材料與人工估驗資料
            DateTime dtStart = DateTime.Parse(Request["dtDate_S"]);
            DateTime dtEnd = DateTime.Parse(Request["dtDate_E"]);
            ViewBag.PR_ID_S = dtStart.ToString("yyyy/MM/dd");
            ViewBag.PR_ID_E = dtEnd.ToString("yyyy/MM/dd");
            DataSet dsMat = service.getContractFromReport(Request["projectId"], dtStart, dtEnd, "N");
            DataSet dsHum = service.getContractFromReport(Request["projectId"], dtStart, dtEnd, "Y");
            List<DataSet> ds = new List<DataSet>();
            ds.Add(dsMat);
            ds.Add(dsHum);
            contract.lstEstFromDailyReport = ds;
            contract.dsTempWorkDailyReport = service.getTempWorkFromDailyReport(Request["projectId"], dtStart, dtEnd);
            return View("Valuation", contract);
        }
        public ActionResult AddEST(string id, string projectid, string type)
        {
            logger.Info("create EST form id process!  Contract id=" + id);
            //新增估驗單編號
            string formid = service.getEstNo();
            return RedirectToAction("ContractItems", "Estimation", new { id = id, formid = formid, projectid = projectid, type = type });
        }
        /// <summary>
        /// 新增估驗單
        /// 1.提供查詢驗收單、選擇其他代付之估驗單等資料
        /// </summary>
        /// <returns></returns>
        public ActionResult createEstimationOrder()
        {
            //傳入驗收單建立估驗單
            string projectId = Request["projectId"];
            string contracId = Request["contractId"];
            string pr_id_s = Request["PR_ID_S"];
            string pr_id_e = Request["PR_ID_E"];
            string isWage = Request["IsWage"];
            EstimationService service = new EstimationService();
            ContractModels contract = null;
            contract = service.getContract(projectId, contracId, pr_id_s, pr_id_e, isWage);

            ViewBag.contracId = contracId;
            ViewBag.prid_s = pr_id_s;
            ViewBag.prid_e = pr_id_e;
            logger.Debug("get estimation Info :" + contract.supplier.COMPANY_NAME);
            return View(contract);
        }
        /// <summary>
        /// 建立點工的估驗單
        /// </summary>
        public ActionResult createEstimationOrder4TmpOrder()
        {
            //傳入驗收單建立估驗單
            string projectId = Request["projectId"];
            string rtpStartId = Request["rptStartId"];
            string rtpEndId = Request["rptEndId"];
            string supplierId = Request["supplierId"];
            string chargeId = Request["chargeId"];
            ViewBag.contracId = Request["contractId"];
            EstimationService service = new EstimationService();
            ContractModels contract = null;
            contract = service.getContract4TempWork(projectId, rtpStartId, rtpEndId, supplierId, chargeId);
            logger.Debug("get estimation Info :" + contract.supplier.COMPANY_NAME);
            return View("createEstimationOrder", contract);
        }
        /// <summary>
        /// 顯示/更新估驗單
        /// </summary>
        public ActionResult createEstimationOrderInfo()
        {
            //取得現有估驗單，產生編輯畫面
            string formid = Request["formid"];
            logger.Debug("get estimation order=" + formid);
            EstimationService service = new EstimationService();
            ContractModels contract = service.getEstimationOrder(formid);
            ViewBag.contracId = contract.planEST.CONTRACT_ID;
            //將合約資料存入Session
            Session["contract"] = contract;
            return View("createEstimationOrder", contract);
        }
        //取得估驗單彙整資訊
        public ActionResult createEstimationOrder_Approve()
        {
            //取得現有估驗單，產生編輯畫面
            string formid = Request["formid"];
            logger.Debug("get estimation order=" + formid);
            EstimationService service = new EstimationService();
            ContractModels contract = service.getEstimationOrder(formid);
            //計算支付金額與相關數據

            ViewBag.contracId = contract.planEST.CONTRACT_ID;
            return View(contract);
        }
        /// <summary>
        /// 儲存估驗單與相關紀錄
        /// </summary>
        /// <returns></returns>
        public string saveEstimationOrder(FormCollection f)
        {
            string s = JsonConvert.SerializeObject(Request.Form).ToString();
            logger.Debug("Request Data=" + s);
            //1.取得估價單查詢條件
            string formId = f["formId"];
            string projectId = f["projectId"];
            string contracId = f["contractId"];
            string pr_id_s = f["prid_s"];
            string pr_id_e = f["prid_e"];
            string supplierId = f["supplierId"];
            string indirect_cost_type = f["indirect_cost_type"];
            string strStatus = f["status"];
            DateTime? paymentDate = null;
            if (f["payment_date"] != "")
            {
                paymentDate = DateTime.Parse(f["payment_date"]);
            }

            //2.建立估驗單物件
            PLAN_ESTIMATION_FORM form = new PLAN_ESTIMATION_FORM();
            if (formId != "")
            {
                form.EST_FORM_ID = formId;
            }
            form.PROJECT_ID = projectId;
            form.CONTRACT_ID = contracId;
            form.INDIRECT_COST_TYPE = indirect_cost_type;
            form.PAYMENT_DATE = paymentDate;
            if (strStatus == "")
            {
                form.STATUS = 0;
            }else
            {
                form.STATUS =  Convert.ToInt16(strStatus);
            }


            ////3.代付廠商資料
            List<PLAN_ESTIMATION_HOLDPAYMENT> lstHoldPayment = null;
            string supplisers = f["hold4Supplier"];
            if (null != supplisers)
            {
                lstHoldPayment = new List<PLAN_ESTIMATION_HOLDPAYMENT>();
                string[] arySupplier = supplisers.Split(',');
                string[] aryAmount = (string[])f.GetValue("holdAmount").RawValue;
                string[] aryRemark = f["hold4Remark"].Split(',');
                for (int i = 0; i < arySupplier.Length; i++)
                {
                    PLAN_ESTIMATION_HOLDPAYMENT hold = new PLAN_ESTIMATION_HOLDPAYMENT();
                    hold.SUPPLIER_ID = arySupplier[i];
                    try
                    {
                        hold.HOLD_AMOUNT = decimal.Parse(aryAmount[i]);
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex.Message + ",ex=" + ex.StackTrace);
                    }
                    hold.REMARK = aryRemark[i];
                    lstHoldPayment.Add(hold);
                }
            }
            //3.1 廠商請款發票明細資料
            List<PLAN_ESTIMATION_INVOICE> lstInvoice = getEstimationInvoice(f, formId, contracId);
            //4.取得代扣款項資料
            List<PLAN_ESTIMATION_PAYMENT_TRANSFER> lstPaymentTransfer = null;
            string strPaymentTransfer = f["HoldformId"];
            string strChkNoPay = f["hf_notpay"];
            logger.Debug("check:" + strPaymentTransfer);
            if (null != strPaymentTransfer && strPaymentTransfer != "")
            {
                lstPaymentTransfer = new List<PLAN_ESTIMATION_PAYMENT_TRANSFER>();
                //憑單編號
                string[] aryEstForm = strPaymentTransfer.Split(',');
                //代付金額
                string[] aryHoldAmount = (string[])f.GetValue("hf_paidamount").RawValue;
                //手續費
                string[] aryFee = (string[])f.GetValue("hf_fee").RawValue;
                //應扣金額
                string[] aryHoldAmountNeed = (string[])f.GetValue("hf_holdAmount").RawValue;
                //本期扣款
                string[] aryCurholdAmount = (string[])f.GetValue("hf_cur_holdAmount").RawValue;
                //原因
                string[] aryRemark = f["hf_remark"].Split(',');
                for (int i = 0; i < aryEstForm.Length; i++)
                {
                    //string idx = strChkNoPay.IndexOf(aryEstForm[i]).ToString();
                    if (null == strChkNoPay || strChkNoPay.IndexOf(aryEstForm[i]) == -1)
                    {
                        PLAN_ESTIMATION_PAYMENT_TRANSFER paymentTrans = new PLAN_ESTIMATION_PAYMENT_TRANSFER();
                        paymentTrans.PAYMENT_FORM_ID = aryEstForm[i];
                        if (aryHoldAmount[i].Trim() != "")
                        {
                            paymentTrans.PAID_AMOUNT = decimal.Parse(aryHoldAmount[i]);  //代付金額
                        }
                        if (aryFee[i].Trim() != "")
                        {
                            paymentTrans.FEE = decimal.Parse(aryFee[i]);//手續費
                        }
                        if (aryHoldAmountNeed[i].Trim() != "")
                        {
                            paymentTrans.HOLD_AMOUNT = decimal.Parse(aryHoldAmountNeed[i]);//應扣金額
                        }
                        if (aryCurholdAmount[i] != "")
                        {
                            paymentTrans.CUR_HOLDAMOUNT = decimal.Parse(aryCurholdAmount[i]);//本期扣款
                        }
                        paymentTrans.REMARK = aryRemark[i];
                        lstPaymentTransfer.Add(paymentTrans);
                    }
                }
            }

            form.PAYEE = supplierId;
            SYS_USER u = UtilService.getUserInfoFromSession(Session);
            form.CREATE_ID = u.USER_ID;
            form.CREATE_DATE = DateTime.Now;

            //5.建立估驗單資料
            EstimationService service = new EstimationService();
            if (formId == "" || null == formId)
            {
                service.createEstimationOrder(form, lstHoldPayment, lstPaymentTransfer, lstInvoice, pr_id_s, pr_id_e);
            }
            else
            {
                service.modifyEstimationOrder(form, lstHoldPayment, lstPaymentTransfer, lstInvoice);
            }


            ///return RedirectToAction("Edit", new { id = 1 });
            logger.Debug("get form Id=" + form.EST_FORM_ID);
            return form.EST_FORM_ID;
        }
        /// <summary>
        /// 取得表單Inviice Records
        /// </summary>
        private static List<PLAN_ESTIMATION_INVOICE> getEstimationInvoice(FormCollection f, string formId, string contracId)
        {
            List<PLAN_ESTIMATION_INVOICE> lstInvoice = null;
            string Invoices = f["invoiceNo"];
            if (null != Invoices && "" != Invoices)
            {
                lstInvoice = EstimationService.getSellInvoice(f, formId, contracId, lstInvoice, Invoices);
            }
            return lstInvoice;
        }

        //1.刪除估驗單資料
        public string delEstimationOrder(FormCollection f)
        {
            string formId = f["formId"];
            SYS_USER u = UtilService.getUserInfoFromSession(Session);
            logger.Info(u.USER_ID + " delete estimation form=" + formId);
            EstimationService service = new EstimationService();
            service.delEstimationOrder(formId);
            return formId;
        }
        public ActionResult ContractItems(string projectid)
        {
            //logger.Info("Access To Contract Item By Contract Id =" + id);
            logger.Info("Access To Est Form By Project Id =" + projectid);
            string formid = service.getEstNo();
            ViewBag.projectId = projectid;
            TND_PROJECT p = service.getProjectById(projectid);
            ViewBag.projectName = p.PROJECT_NAME;
            //ViewBag.wage = type;
            ExpenseTask contract = new ExpenseTask();
            ViewBag.formid = formid;
            ViewBag.date = DateTime.Now;
            //ViewBag.contractid = id;
            //ViewBag.paymentTermsId = id + '/' + formid;
            //ViewBag.keyid = id; //使用供應商名稱的contractid
            //取得合約金額與供應商名稱,採購項目等資料
            /*
            if (ViewBag.wage != "Y")
            {
                plansummary lstContract = service.getPlanContract4Est(id);
                ViewBag.supplier = lstContract.SUPPLIER_ID;
                ViewBag.formname = lstContract.FORM_NAME;
                ViewBag.amount = lstContract.MATERIAL_COST;
                //ViewBag.contractid = lstContract.CONTRACT_ID; //使用供應商編號的contractid
                PaymentTermsFunction payment = service.getPaymentTerm(id, formid);
                if (payment.PAYMENT_RETENTION_RATIO != null)
                {
                    ViewBag.retention = payment.PAYMENT_RETENTION_RATIO;
                }
                else if (payment.USANCE_RETENTION_RATIO != null)
                {
                    ViewBag.retention = payment.USANCE_RETENTION_RATIO;
                }
                if (payment.PAYMENT_TYPE == "連工帶料")
                {
                    ViewBag.type = 'C'; // 合約包含材料與工資
                }
                else
                {
                    ViewBag.type = 'M'; // 設備合約
                }
                ViewBag.estCount = service.getEstCountById(id);
                ViewBag.paymentTerms = service.getTermsByContractId(id);
                var balance = service.getBalanceOfRefundById(id);
                if (balance > 0)
                {
                    TempData["balance"] = "本合約目前尚有 " + string.Format("{0:C0}", balance) + "的代付支出款項，仍未扣回!";
                }
            }
            else
            {
                plansummary lstWageContract = service.getPlanContractOfWage4Est(id);
                ViewBag.supplier4Wage = lstWageContract.MAN_SUPPLIER_ID;
                ViewBag.formname4Wage = lstWageContract.MAN_FORM_NAME;
                ViewBag.amount4Wage = lstWageContract.WAGE_COST;
                ViewBag.contractid4Wage = lstWageContract.CONTRACT_ID;
                ViewBag.type = 'W'; //工資合約
                PaymentTermsFunction payment = service.getPaymentTerm(lstWageContract.CONTRACT_ID, formid);
                if (payment.PAYMENT_RETENTION_RATIO != null)
                {
                    ViewBag.retention = payment.PAYMENT_RETENTION_RATIO;
                }
                else if (payment.USANCE_RETENTION_RATIO != null)
                {
                    ViewBag.retention = payment.USANCE_RETENTION_RATIO;
                }
                ViewBag.estCount = service.getEstCountById(lstWageContract.CONTRACT_ID);
                ViewBag.paymentTerms = service.getTermsByContractId(lstWageContract.CONTRACT_ID);
                var balance = service.getBalanceOfRefundById(lstWageContract.CONTRACT_ID);
                if (balance > 0)
                {
                    TempData["balance"] = "本合約目前尚有 " + string.Format("{0:C0}", balance) + "的代付支出款項，仍未扣回!";
                }
            }
            //ViewBag.paymentkey = ViewBag.formid + ViewBag.contractid;
            ViewBag.InvoicePieces = service.getInvoicePiecesById(formid);
            List<EstimationForm> lstContractItem = null;
            lstContractItem = service.getContractItemById(id, projectid);
            //contract.planItems = lstContractItem;
            ViewBag.SearchResult = "共取得" + lstContractItem.Count + "筆資料";
            //轉成Json字串
            ViewData["items"] = JsonConvert.SerializeObject(lstContractItem);
            */

            return View("ContractItems", contract);
        }

        //DOM 估驗作業功能紐對應不同Action
        public class MultiButtonAttribute : ActionNameSelectorAttribute
        {
            public string Name { get; set; }
            public MultiButtonAttribute(string name)
            {
                this.Name = name;
            }
            public override bool IsValidName(ControllerContext controllerContext,
                string actionName, System.Reflection.MethodInfo methodInfo)
            {
                if (string.IsNullOrEmpty(this.Name))
                {
                    return false;
                }
                return controllerContext.HttpContext.Request.Form.AllKeys.Contains(this.Name);
            }
        }
        //啟動申請程序)
        public string createFlow(FormCollection f)
        {
            //取得估驗單編號
            logger.Info("EST form Id:" + Request["formid"]);
            SYS_USER u = (SYS_USER)Session["user"];
            logger.Info("update Estimation Form");
            //估驗單草稿 STATUS = 0
            int k = service.UpdateESTStatusById(Request["formid"]);
            Flow4EstimationForm flowService = new Flow4EstimationForm();
            flowService.iniRequest(u, Request["formid"]);
            return "測試!!";//RedirectToAction("SingleEST", "Estimation", new { id = Request["formid"] });
            //}
        }
        //審核估驗單
        public string SendEstimationForm(FormCollection f)
        {
            logger.Info("http get mehtod:" + f["EST_FORM_ID"]);
            //取得單據與任務資料
            Flow4EstimationForm wfs = new Flow4EstimationForm();
            wfs.getTask(f["formId"]);
            //廠商請款發票明細資料，提供業管人員於審核時輸入
            List<PLAN_ESTIMATION_INVOICE> lstInvoice = getEstimationInvoice(f, f["formId"], f["contracId"]);
            if (null == lstInvoice || lstInvoice.Count == 0)
            {
                Response.StatusCode = 500;//Anything other than 2XX HTTP status codes should work
                Response.Write("付款憑證發票未輸入!!");
                return null;
            }
            SYS_USER u = (SYS_USER)Session["user"];
            DateTime? date = null;//DateTime can not set null
            string desc = null;
            string payee = null;
            string remark = null;
            if (f["payment_date"].ToString() != "")
            {
                date = Convert.ToDateTime(f["payment_date"].ToString());
            }
            if (null != f["RejectDesc"] && f["RejectDesc"].ToString() != "")
            {
                desc = f["RejectDesc"].ToString().Trim();
            }
            if (null != f["supplierId"] && f["supplierId"].ToString() != "")
            {
                payee = f["supplierId"].ToString().Trim();
            }
            if (null != f["remark"] && f["remark"].ToString() != "")
            {
                remark = f["remark"].ToString().Trim();
            }
            wfs.Send(u, date, desc, payee, remark);

            return "更新成功!!";
        }
        //中止估驗單
        public string CancelEstimationForm(FormCollection f)
        {
            Flow4EstimationForm wfs = new Flow4EstimationForm();
            wfs.getTask(f["formId"]);
            logger.Info("Cancel EstimationForm:" + wfs.task.task.EST_FORM_ID);

            SYS_USER u = (SYS_USER)Session["user"];
            wfs.Cancel(u);
            return wfs.Message;
        }
        //退件-估驗單
        public string RejectEstimationForm(FormCollection f)
        {
            Flow4EstimationForm wfs = new Flow4EstimationForm();
            wfs.getTask(f["formId"]);
            logger.Info("Reject EstimationForm:" + wfs.task.task.EST_FORM_ID);
            SYS_USER u = (SYS_USER)Session["user"];
            DateTime? payDate = null;
            if (f["payment_date"] != "")
            {
                payDate = DateTime.Parse(f["payment_date"]);
            }
            wfs.Reject(u, payDate, f["RejectDesc"], f["supplierId"]);
            return wfs.Message;
        }
        /* 分隔線- ----先前功能未考慮流程須檢視  -------------------------*/
        [HttpPost]
        [MultiButton("AddEst")]
        //驗收單送審
        public ActionResult AddEst(FormCollection f)
        {
            //取得估驗單編號
            logger.Info("EST form Id:" + Request["formid"]);
            logger.Info("get EST No of un Approval By contractid:" + Request["contractid"]);
            logger.Info("Project Id:" + Request["projectid"]);
            string contractid = Request["contractid"];
            string UnApproval = null;
            UnApproval = service.getEstNoByContractId(contractid);
            if (UnApproval != null && "" != UnApproval)
            {
                TempData["result"] = "目前尚有未核准的估驗單，估驗單編號為" + UnApproval + "，待此單核准後再新增估驗單!";
                return RedirectToAction("Valuation", "Estimation", new { id = Request["projectid"] });
            }
            else
            {
                //更新估驗單
                logger.Info("update Estimation Form");
                //估驗單(已送審) STATUS = 20
                //int k = service.RefreshESTStatusById(Request["formid"]);
                return RedirectToAction("SingleEST", "Estimation", new { id = Request["formid"] });
            }
        }
        //廠商請款作業
        public String ConfirmEst(PLAN_ESTIMATION_FORM est, HttpPostedFileBase file)
        {
            //取得專案編號
            logger.Info("Project Id:" + Request["id"]);
            //建立估驗單
            logger.Info("create new Estimation Form");
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            est.PROJECT_ID = Request["id"];
            est.CREATE_ID = u.USER_ID;
            est.CREATE_DATE = DateTime.Now;
            est.EST_FORM_ID = Request["estid"];
            est.PROJECT_NAME = Request["projectName"];
            est.PAYEE = Request["supplier"];
            if (Request["paymentDate"] != "")
            {
                est.PAYMENT_DATE = Convert.ToDateTime(Request["paymentDate"]);
            }
            //NDIRECT_COST_TYPE : M 代表界面維保費用；O 代表其他(泛指非直接成本與維保費的其他額外延伸的成本)
            if (null != Request["indirect_cost_type"] && Request["indirect_cost_type"] != "")
            {
                est.INDIRECT_COST_TYPE = Request["indirect_cost_type"];
            }
            //est.CONTRACT_ID = Request["contractid"];
            //est.PLUS_TAX = Request["tax"];
            est.REMARK = Request["remark"];
            //est.INVOICE = Request["invoice"];
            est.STATUS = -10;
            try
            {
                est.PAYMENT_TRANSFER = decimal.Parse(Request["paid_amount"]);
            }
            catch (Exception ex)
            {
                logger.Error(est.PAYMENT_TRANSFER + " not payment_transfer:" + ex.Message);
            }
            try
            {
                est.PAID_AMOUNT = decimal.Parse(Request["paid_amount"]);
            }
            catch (Exception ex)
            {
                logger.Error(est.PAID_AMOUNT + " not paid_amount:" + ex.Message);
            }
            //若使用者有上傳計價附檔，則增加檔案資料
            if (null != file && file.ContentLength != 0)
            {
                //TND_FILE saveF = new TND_FILE();
                //2.解析檔案
                logger.Info("Parser file data:" + file.FileName);
                //2.1 設定檔案名稱,實體位址,附檔名
                string projectid = Request["id"];
                string keyName = Request["estid"];
                logger.Info("file upload namme =" + keyName);
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + Request["id"], fileName);
                //saveF.FILE_ACTURE_NAME = fileName;
                var fileType = Path.GetExtension(file.FileName);
                //f.FILE_LOCATIOM = path;
                string createDate = DateTime.Now.ToString("yyyy/MM/dd");
                logger.Info("createDate = " + createDate);
                string createId = uInfo.USER_ID;
                FileManage fs = new FileManage();
                string k = fs.addFile(projectid, keyName, fileName, fileType, path, createId, createDate);
                //int j = service.refreshVAFile(saveF);
                //2.2 將上傳檔案存檔
                logger.Info("save upload file:" + path);
                file.SaveAs(path);
            }
            //string estid = service.newEST(Request["formid"], est, lstItemId);
            string estid = service.newEST(Request["estid"], est);

            int iR = service.UpdateESTStatusById(estid);
            Flow4Estimation flowService = new Flow4Estimation();
            flowService.iniRequest(uInfo, estid);

            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            //string itemJson = objSerializer.Serialize(service.getDetailsPayById(Request["formid"], Request["contractid"]));
            string itemJson = objSerializer.Serialize(service.getDetailsPayById(Request["estid"], Request["estid"]));
            logger.Info("EST form details payment amount=" + itemJson);
            return itemJson;
            //}
        }
        //估驗單查詢
        public ActionResult EstimationForm(string id)
        {
            logger.Info("Search For Estimation Form !!");
            ViewBag.projectid = id;
            getProject(id);
            //取得表單狀態參考資料
            SelectList LstStatus = new SelectList(SystemParameter.getSystemPara("EstimationForm"), "KEY_FIELD", "VALUE_FIELD");
            ViewData.Add("status", LstStatus);
            string strStatus = null;
            if (null != Request["status"])
            {
                strStatus = Request["status"];
            }
            EstimationFormApprove model = new EstimationFormApprove();
            //取得審核中表單資料
            Flow4EstimationForm workflow = new Flow4EstimationForm();
            EstimationService es = new EstimationService();
            model.lstEstimationFlowTask = workflow.getEstimationFormRequest(Request["contractid"], Request["payee"], Request["estid"], id, strStatus);
            //初稿的估驗單
            model.lstEstimationForm = es.getFormList(id, "0");
            ViewBag.SearchResult = "共取得" + model.lstEstimationFlowTask + "筆資料";
            getProject(id);

            return View(model);
        }
        //取得專案資料
        protected void getProject(string id)
        {
            TnderProjectService tndservice = new TnderProjectService();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.projectid = p.PROJECT_ID;
        }

        //取得合約付款條件
        public string getPaymentTerms(string contractid)
        {
            string[] key = Request["contractid"].Split('/');
            PurchaseFormService service = new PurchaseFormService();
            logger.Info("access the terms of payment by:" + Request["contractid"]);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(service.getPaymentTerm(key[0], key[1]));
            logger.Info("plan payment terms info=" + itemJson);
            return itemJson;
        }

        //顯示單一估驗單功能
        //有Double Submit 的問題
        public ActionResult SingleEST(string id)
        {
            logger.Info("http get mehtod:" + id);
            if (null == id)
            {
                id = Request["formid"];
            }
            ContractModels singleForm = new ContractModels();
            Flow4Estimation wfs = new Flow4Estimation();
            service.getESTByEstId(id);
            singleForm.planEST = service.formEST;
            ViewBag.formid = id;

            singleForm.planESTItem = service.ESTItem;
            singleForm.project = service.getProjectById(singleForm.planEST.PROJECT_ID);
            logger.Debug("Project ID:" + singleForm.project.PROJECT_ID);
            //PaymentDetailsFunction lstSummary = service.getDetailsPayById(id, singleForm.planEST.CONTRACT_ID);
            PaymentDetailsFunction lstSummary = service.getDetailsPayById(id, id);
            PaymentDetailsFunction pay = service.getDetailsPayById(id, id);
            ViewBag.loanAmount = 0;
            if (pay.LOAN_AMOUNT != 0)
            {
                TempData["loanAmt"] = "此廠商有借款尚未償還，金額為:";
                ViewBag.loanAmount = pay.LOAN_AMOUNT;
                ViewBag.loanPayee = pay.LOAN_PAYEE_ID;
            }
            //var balance = service.getBalanceOfRefundById(singleForm.planEST.CONTRACT_ID);
            //if (balance > 0)
            //{
            //TempData["balance"] = "本合約目前尚有 " + string.Format("{0:C0}", balance) + "的代付支出款項，仍未扣回!";
            //}
            //轉成Json字串
            //ViewData["items"] = JsonConvert.SerializeObject(singleForm.planESTItem);
            ViewData["summary"] = JsonConvert.SerializeObject(lstSummary);

            logger.Info("get process request by dataId=" + id);
            wfs.getTask(id);
            wfs.getRequest(id);
            wfs.task.EstData = singleForm;

            Session["process"] = wfs.task;
            return View(wfs.task);
        }

        //送審、通過
        public String SendForm(FormCollection f)
        {
            logger.Info("http get mehtod:" + f["EST_FORM_ID"]);
            Flow4Estimation wfs = new Flow4Estimation();
            wfs.task = (ExpenseTask)Session["process"];
            logger.Info("Data In Session :" + wfs.task.EstData.planEST.EST_FORM_ID);

            SYS_USER u = (SYS_USER)Session["user"];
            DateTime? date = null;//DateTime can not set null
            string desc = null;
            string payee = null;
            string remark = null;
            if (f["paymentDate"].ToString() != "")
            {
                date = Convert.ToDateTime(f["paymentDate"].ToString());
            }
            if (null != f["RejectDesc"] && f["RejectDesc"].ToString() != "")
            {
                desc = f["RejectDesc"].ToString().Trim();
            }
            if (null != f["supplier"] && f["supplier"].ToString() != "")
            {
                payee = f["supplier"].ToString().Trim();
            }
            if (null != f["remark"] && f["remark"].ToString() != "")
            {
                remark = f["remark"].ToString().Trim();
            }
            wfs.Send(u, date, desc, payee, remark);

            return "更新成功!!";
        }
        //退件
        public String RejectForm(FormCollection form)
        {
            //取得表單資料 from Session
            Flow4Estimation wfs = new Flow4Estimation();
            wfs.task = (ExpenseTask)Session["process"];
            SYS_USER u = (SYS_USER)Session["user"];
            wfs.Reject(u, form["RejectDesc"]);
            return wfs.Message;
        }
        //取消
        public String CancelForm(FormCollection form)
        {
            Flow4Estimation wfs = new Flow4Estimation();
            wfs.task = (ExpenseTask)Session["process"];
            SYS_USER u = (SYS_USER)Session["user"];
            wfs.Cancel(u);
            return wfs.Message;
        }
        //其他扣款
        public ActionResult OtherPayment(string id, string contractid)
        {
            logger.Info("Access To Other Payment By EST Form Id =" + id);
            service.getInqueryForm(contractid);
            PLAN_SUP_INQUIRY f = service.formInquiry;
            ViewBag.projectId = f.PROJECT_ID;
            TND_PROJECT p = service.getProjectById(f.PROJECT_ID);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.contractid = contractid;
            ViewBag.formname = f.FORM_NAME;
            ViewBag.formid = id;
            List<PLAN_OTHER_PAYMENT> lstOtherPayItem = null;
            lstOtherPayItem = service.getOtherPayById(id);
            ViewBag.status = -10; //估驗單尚未建立
            ViewBag.status = service.getStatusById(id);
            ViewBag.key = lstOtherPayItem.Count;
            logger.Debug("this other payment record =" + ViewBag.key + "筆");
            ViewData["items"] = JsonConvert.SerializeObject(lstOtherPayItem);
            return View();
        }

        //更新其他扣款
        public String UpdateOtherPay(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            // 先刪除原先資料
            logger.Info("EST form id =" + form["formid"]);
            logger.Info("Delete PLAN_OTHER_PAYMENT By EST_FORM_ID");
            service.delOtherPayByESTId(form["formid"]);
            // 再次取得其他扣款資料
            string[] lstAmount = form.Get("input_amount").Split(',');
            string[] lstReason = form.Get("input_reason").Split(',');
            List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
            for (int j = 0; j < lstAmount.Count(); j++)
            {
                PLAN_OTHER_PAYMENT item = new PLAN_OTHER_PAYMENT();
                item.EST_FORM_ID = form["formid"];
                item.CONTRACT_ID = form["contractid"];
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                logger.Info("Other Payment Amount  =" + item.AMOUNT);
                item.REASON = lstReason[j];
                logger.Debug("Item EST form id =" + item.EST_FORM_ID + "且扣款原因為" + item.REASON);
                lstItem.Add(item);
            }
            int i = service.addOtherPayment(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新其他扣款資料成功，EST_FORM_ID =" + form["formid"];
            }
            return msg;
        }
        /*
        //更新估驗數量與稅率
        public String UpdateESTQty(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            string tax = form.Get("tax").Trim();
            decimal taxratio = 0;
            if (tax == "E")
            {
                taxratio = decimal.Parse(form.Get("taxratio").Trim());
            }
            string[] lstPlanItemId = form.Get("planitemid").Split(',');
            string[] lstQty = form.Get("evaluated_qty").Split(',');
            List<PLAN_ESTIMATION_ITEM> lstItem = new List<PLAN_ESTIMATION_ITEM>();
            for (int j = 0; j < lstPlanItemId.Count(); j++)
            {
                PLAN_ESTIMATION_ITEM item = new PLAN_ESTIMATION_ITEM();
                item.PLAN_ITEM_ID = lstPlanItemId[j];
                if (lstQty[j].ToString() == "")
                {
                    item.EST_QTY = null;
                }
                else
                {
                    item.EST_QTY = decimal.Parse(lstQty[j]);
                }
                logger.Debug("EST_FIRM_ID = " + form["estid"] + "It's Plan tem Id No=" + item.PLAN_ITEM_ID + ", EST Qty =" + item.EST_QTY);
                lstItem.Add(item);
            }
            int k = service.RefreshESTByEstId(form["estid"], tax, taxratio);
            int i = service.refreshESTQty(form["estid"], lstItem);
            int m = service.UpdateRetentionAmountById(form["estid"], form["contractid"]);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新估驗數量成功，請確認各項金額是否正確 !";
            }

            logger.Info("Request: 更新數量訊息 = " + msg);
            return msg;
        }
        */
        //更新估驗單
        public String UpdateEST(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            /*
            decimal foreign_payment = 0;

            if (null != form["t_foreign"] && form["t_foreign"] != "")
            {
                foreign_payment = decimal.Parse(form.Get("t_foreign").Trim());
            }
            decimal original_foreign_payment = decimal.Parse(form.Get("original_t_foreign").Trim());
            decimal retention = 0;
            if (null != form["t_retention"] && form["t_retention"] != "")
            {
                retention = decimal.Parse(form.Get("t_retention").Trim());
            }
            decimal subAmount = decimal.Parse(form.Get("sub_amount").Trim());
            decimal repayment = decimal.Parse(form.Get("t_repayment").Trim());
            //decimal totalAmount = decimal.Parse(form.Get("totalAmount").Trim());
            decimal tax_ratio = decimal.Parse(form.Get("taxratio").Trim());
            decimal tax_amount = 0;
            if (foreign_payment != original_foreign_payment)
            {
                tax_amount = Math.Round((subAmount - repayment) * tax_ratio / 100, 0);
            }
            else
            {
                tax_amount = decimal.Parse(form.Get("tax_amount").Trim());
            }
            */
            decimal paidAmount = decimal.Parse(form.Get("paid_amount").Trim());
            string remark = form.Get("remark").Trim();
            string payee = form.Get("supplier").Trim();
            string projectName = form.Get("projectName").Trim();
            string indirectCostType = "";
            if (null != form.Get("indirect_cost_type").Trim() && form.Get("indirect_cost_type").Trim() != "")
            {
                indirectCostType = form.Get("indirect_cost_type").Trim();
            }
            DateTime? paymentDate = null;
            if (form.Get("paymentDate").Trim() != "")
            {
                paymentDate = Convert.ToDateTime(form.Get("paymentDate").Trim());
            }
            //int i = service.RefreshESTAmountByEstId(form["estid"], subAmount, foreign_payment, retention, tax_amount, remark);
            //int i = service.RefreshESTAmountByEstId(form["estid"], paidAmount, 0, 0, 0, remark);
            int i = service.RefreshESTAmountByEstId(form["estid"], paidAmount, payee, projectName, paymentDate, remark, indirectCostType);
            //修改小計金額(PAYMENT_TRANSFER 欄位)與實付金額(PAID_AMOUNT 欄位)
            //PaymentDetailsFunction amountPaid = service.getDetailsPayById(form["estid"], form["contractid"]);
            int t = service.UpdatePaidAmountById(form["estid"], paidAmount);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新估驗單成功";
            }

            logger.Info("Request: 更新估驗單訊息 = " + msg);
            return msg;
        }

        //預付款
        public ActionResult AdvancePayment(string id, string contractid)
        {
            logger.Info("Access To Advance Payment By EST Form Id =" + id);
            ContractModels contract = getContractFromSession(id, contractid);
            PaymentTermsFunction payment = contract.contractPaymentTerms;
            ViewBag.advancePaymentRatio = payment.PAYMENT_ADVANCE_RATIO;
            AdvancePaymentFunction advancePay = service.getAdvancePayById(id, contractid);
            List<PLAN_OTHER_PAYMENT> lstAdvancePayItem = null;
            lstAdvancePayItem = service.getAdvancePayByESTId(id);
            ViewBag.key = lstAdvancePayItem.Count;
            ViewData["items"] = JsonConvert.SerializeObject(advancePay);
            return View();
        }
        // 取得預付款資料
        public String AddAdvancePay(FormCollection form)
        {
            //A:預付款, B:暫借款 C: 保證金
            logger.Info("form:" + form.Count);
            string msg = "";
            string advance_payment = form.Get("advance_payment");
            string temporary_loan = form.Get("temporary_loan");
            string margins = form.Get("margins");
            string[] lstItemId = String.Join(",", advance_payment, temporary_loan, margins).Split(',');
            string[] lsttype = String.Join(",", "A", "B", "C").Split(',');

            List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_OTHER_PAYMENT item = new PLAN_OTHER_PAYMENT();
                item.EST_FORM_ID = form["formid"];
                item.CONTRACT_ID = form["contractid"];
                if (lstItemId[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstItemId[j]);
                }
                logger.Info("Advance Payment Amount  =" + item.AMOUNT);
                item.TYPE = lsttype[j];
                logger.Debug("Item EST form id =" + item.EST_FORM_ID + " ,And Type =" + item.TYPE);
                lstItem.Add(item);
            }
            int i = service.addAdvancePayment(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "新增預付款資料成功，EST_FORM_ID =" + form["formid"];
            }
            return msg;
        }
        //更新預付款資料
        public String UpdateAdvancePay(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            // 取得預付款資料
            string advance_payment = form.Get("advance_payment");
            string temporary_loan = form.Get("temporary_loan");
            string margins = form.Get("margins");
            string[] lstItemId = String.Join(",", advance_payment, temporary_loan, margins).Split(',');
            string[] lsttype = String.Join(",", "A", "B", "C").Split(',');
            List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_OTHER_PAYMENT item = new PLAN_OTHER_PAYMENT();
                item.EST_FORM_ID = form["formid"];
                if (lstItemId[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstItemId[j]);
                }
                logger.Info("Advance Payment Amount  =" + item.AMOUNT);
                item.TYPE = lsttype[j];
                logger.Debug("Item EST form id =" + item.EST_FORM_ID + " ,And Type =" + item.TYPE);
                lstItem.Add(item);
            }
            int i = service.updateAdvancePayment(form["formid"], lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "修改預付款資料成功，EST_FORM_ID =" + form["formid"];
            }
            return msg;
        }
        /*
       public String UpdateESTStatusById(FormCollection form)
       {
           //取得估驗單編號
           logger.Info("form:" + form.Count);
           logger.Info("EST form Id:" + form["estid"]);
           string msg = "";
           decimal foreign_payment = decimal.Parse(form.Get("t_foreign").Trim());
           decimal original_foreign_payment = decimal.Parse(form.Get("original_t_foreign").Trim());
           decimal retention = decimal.Parse(form.Get("t_retention").Trim());
           //decimal totalAmount = decimal.Parse(form.Get("totalAmount").Trim());
           decimal subAmount = decimal.Parse(form.Get("sub_amount").Trim());
           decimal repayment = decimal.Parse(form.Get("t_repayment").Trim());
           decimal tax_ratio = decimal.Parse(form.Get("taxratio").Trim());
           decimal tax_amount = 0;
           if (foreign_payment != original_foreign_payment)
           {
               tax_amount = Math.Round((subAmount - repayment) * tax_ratio / 100, 0);
           }
           else
           {
               tax_amount = decimal.Parse(form.Get("tax_amount").Trim());
           }
           string remark = form.Get("remark").Trim();
           int i = service.RefreshESTAmountByEstId(form["estid"], subAmount, foreign_payment, retention, tax_amount, remark);
           //修改小計金額(PAYMENT_TRANSFER 欄位)與實付金額(PAID_AMOUNT 欄位)
           PaymentDetailsFunction amountPaid = service.getDetailsPayById(form["estid"], form["contractid"]);
           int t = service.UpdatePaidAmountById(form["estid"], decimal.Parse(amountPaid.PAID_AMOUNT.ToString()));
           //更新估驗單狀態
           logger.Info("Update Estimation Form Status");
           //估驗單(已送審) STATUS = 20
           int j = service.RefreshESTStatusById(form["estid"]);
           if (j == 0)
           {
               msg = service.message;
           }
           else
           {
               msg = "估驗單已送審";
           }
           return msg;
       }


       public String RejectESTById(FormCollection form)
       {
           //取得估驗單編號
           logger.Info("EST form Id:" + form["estid"]);
           //更新估驗單狀態
           logger.Info("Reject Estimation Form ");
           //估驗單(已退件) STATUS = 0
           string msg = "";
           int i = service.RejectESTByEstId(form["estid"]);
           if (i == 0)
           {
               msg = service.message;
           }
           else
           {
               msg = "估驗單已退回";
           }
           return msg;
       }

       public String ApproveESTById(FormCollection form)
       {
           //取得估驗單編號
           logger.Info("EST form Id:" + form["estid"]);
           //更新估驗單狀態
           logger.Info("Confirm Estimation Form ");
           //估驗單(已核可) STATUS = 30
           string msg = "";
           int i = service.ApproveESTByEstId(form["estid"]);
           if (i == 0)
           {
               msg = service.message;
           }
           else
           {
               msg = "估驗單已核可";
           }
           return msg;
       }
       */
        //憑證
        public ActionResult Invoice(string id, string contractid)
        {
            logger.Info("Access To Invoice By EST Form Id =" + id);
            service.getInqueryForm(contractid);
            PLAN_SUP_INQUIRY f = service.formInquiry;
            ViewBag.projectId = f.PROJECT_ID;
            TND_PROJECT p = service.getProjectById(f.PROJECT_ID);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.formname = f.FORM_NAME;
            ViewBag.contractid = contractid;
            ViewBag.formid = id;
            TND_SUPPLIER lstSupplier = service.getSupplierInfo(f.SUPPLIER_ID);
            ViewBag.supplier = lstSupplier.COMPANY_NAME;
            ViewBag.companyid = lstSupplier.COMPANY_ID;
            ContractModels singleForm = new ContractModels();
            int EstCount = service.getEstCountById(id);
            ViewBag.type = "不檢查發票";
            ViewBag.tax = "其他";
            if (EstCount > 1)
            {
                service.getESTByEstId(id);
                singleForm.planEST = service.formEST;
                try
                {
                    if (singleForm.planEST.INVOICE == "E")
                    {
                        ViewBag.type = "發票含保留款";
                    }
                    else if (singleForm.planEST.INVOICE == "I")
                    {
                        ViewBag.type = "發票不含保留款";
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex.StackTrace);
                }
                try
                {
                    if (singleForm.planEST.PLUS_TAX == "E")
                    {
                        ViewBag.tax = "外加稅 " + singleForm.planEST.TAX_RATIO + "  %";
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex.StackTrace);
                }
            }
            List<PLAN_INVOICE> lstInvoice = null;
            List<PLAN_INVOICE> allInvoice = null;
            lstInvoice = service.getInvoiceById(id);
            ViewBag.key = lstInvoice.Count;
            logger.Debug("this invoice record =" + ViewBag.key + "筆");
            ViewData["items"] = JsonConvert.SerializeObject(lstInvoice);
            allInvoice = service.getInvoiceByContractId(id, contractid);
            return View(allInvoice);
        }

        public String AddInvoice(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            // 取得憑證資料
            string[] lstDate = form.Get("invoice_date").Split(',');
            string[] lstNumber = form.Get("invoice_number").Split(',');
            string[] lstAmount = form.Get("input_amount").Split(',');
            string[] lstTax = form.Get("taxamount").Split(',');
            string[] lstType = form.Get("invoicetype").Split(',');
            string[] lstSubType = form.Get("sub_type").Split(',');
            string[] lstPlanItem = form.Get("plan_item_id").Split(',');
            string[] lstDiscountQty = form.Get("discount_qty").Split(',');
            string[] lstDiscountPrice = form.Get("discount_unit_price").Split(',');
            List<PLAN_INVOICE> lstItem = new List<PLAN_INVOICE>();
            for (int j = 0; j < lstAmount.Count(); j++)
            {
                PLAN_INVOICE item = new PLAN_INVOICE();
                item.EST_FORM_ID = form["formid"];
                item.CONTRACT_ID = form["contractid"];
                item.INVOICE_NUMBER = lstNumber[j];
                try
                {
                    item.INVOICE_DATE = Convert.ToDateTime(lstDate[j]);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.StackTrace);
                }
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                if (lstTax[j].ToString() == "")
                {
                    item.TAX = null;
                }
                else
                {
                    item.TAX = decimal.Parse(lstTax[j]);
                }
                if (lstType[j].ToString() == "")
                {
                    item.TYPE = null;
                }
                else
                {
                    item.TYPE = lstType[j];
                }
                if (lstSubType[j].ToString() == "")
                {
                    item.SUB_TYPE = null;
                }
                else
                {
                    item.SUB_TYPE = lstSubType[j];
                }
                item.PLAN_ITEM_ID = lstPlanItem[j];
                if (lstDiscountQty[j].ToString() == "")
                {
                    item.DISCOUNT_QTY = null;
                }
                else
                {
                    item.DISCOUNT_QTY = decimal.Parse(lstDiscountQty[j]);
                }
                if (lstDiscountPrice[j].ToString() == "")
                {
                    item.DISCOUNT_UNIT_PRICE = null;
                }
                else
                {
                    item.DISCOUNT_UNIT_PRICE = decimal.Parse(lstDiscountPrice[j]);
                }
                logger.Info("Invoice Number = " + item.INVOICE_NUMBER + "and Invoice Amount =" + item.AMOUNT);
                //logger.Debug("Item EST form id =" + item.EST_FORM_ID + "且憑證類型為" + item.TYPE);
                lstItem.Add(item);
            }
            int i = service.addInvoice(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "新增憑證資料成功，EST_FORM_ID =" + form["formid"];
            }
            return msg;
        }
        //更新憑證資料??
        public String UpdateInvoice(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            // 先刪除原先資料
            logger.Info("EST form id =" + form["formid"]);
            logger.Info("Delete PLAN_INVOICE By EST_FORM_ID");
            service.delInvoiceByESTId(form["formid"]);
            // 再次取得憑證資料
            string[] lstDate = form.Get("invoice_date").Split(',');
            string[] lstNumber = form.Get("invoice_number").Split(',');
            string[] lstAmount = form.Get("input_amount").Split(',');
            string[] lstTax = form.Get("taxamount").Split(',');
            string[] lstType = form.Get("invoicetype").Split(',');
            string[] lstSubType = form.Get("sub_type").Split(',');
            string[] lstPlanItem = form.Get("plan_item_id").Split(',');
            string[] lstDiscountQty = form.Get("discount_qty").Split(',');
            string[] lstDiscountPrice = form.Get("discount_unit_price").Split(',');
            List<PLAN_INVOICE> lstItem = new List<PLAN_INVOICE>();
            for (int j = 0; j < lstAmount.Count(); j++)
            {
                PLAN_INVOICE item = new PLAN_INVOICE();
                item.EST_FORM_ID = form["formid"];
                item.CONTRACT_ID = form["contractid"];
                item.INVOICE_NUMBER = lstNumber[j];
                try
                {
                    item.INVOICE_DATE = Convert.ToDateTime(lstDate[j]);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.StackTrace);
                }
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                if (lstTax[j].ToString() == "")
                {
                    item.TAX = null;
                }
                else
                {
                    item.TAX = decimal.Parse(lstTax[j]);
                }
                if (lstType[j].ToString() == "")
                {
                    item.TYPE = null;
                }
                else
                {
                    item.TYPE = lstType[j];
                }
                if (lstType[j] == "折讓單")
                {
                    if (lstSubType[j].ToString() == "")
                    {
                        item.SUB_TYPE = null;
                    }
                    else
                    {
                        item.SUB_TYPE = lstSubType[j];
                    }
                    item.PLAN_ITEM_ID = lstPlanItem[j];
                    if (lstDiscountQty[j].ToString() == "")
                    {
                        item.DISCOUNT_QTY = null;
                    }
                    else
                    {
                        item.DISCOUNT_QTY = decimal.Parse(lstDiscountQty[j]);
                    }
                    if (lstDiscountPrice[j].ToString() == "")
                    {
                        item.DISCOUNT_UNIT_PRICE = null;
                    }
                    else
                    {
                        item.DISCOUNT_UNIT_PRICE = decimal.Parse(lstDiscountPrice[j]);
                    }
                }
                else
                {
                    item.SUB_TYPE = null;
                    item.PLAN_ITEM_ID = null;
                    item.DISCOUNT_QTY = null;
                    item.DISCOUNT_UNIT_PRICE = null;
                }
                logger.Info("Invoice Number = " + item.INVOICE_NUMBER + "and Invoice Amount =" + item.AMOUNT);
                lstItem.Add(item);
            }
            int i = service.addInvoice(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新憑證資料成功，EST_FORM_ID =" + form["formid"];
            }
            return msg;
        }

        //代付支出
        public ActionResult RePayment(string id, string contractid)
        {
            logger.Info("Access To RePayment By EST Form Id =" + id);
            ContractModels contract = getContractFromSession(id, contractid);
            ViewBag.key = contract.EstimationHoldPayments.Count();
            logger.Debug("this repayment record =" + ViewBag.key + "筆");
            ViewData["items"] = JsonConvert.SerializeObject(contract.EstimationHoldPayments);
            logger.Debug(ViewData["items"]);
            return View();
        }

        private ContractModels getContractFromSession(string id, string contractid)
        {
            //將取得Session Data
            ContractModels contract = (ContractModels)Session["contract"];

            service.getInqueryForm(contractid);
            //PLAN_SUP_INQUIRY f = service.formInquiry;
            ViewBag.projectId = contract.project.PROJECT_ID;
            ViewBag.projectName = contract.project.PROJECT_NAME;
            ViewBag.contractid = contract.planEST.CONTRACT_ID;
            ViewBag.formname = contract.supContract.FORM_NAME;
            ViewBag.formid = id;
            ViewBag.status = contract.planEST.STATUS; //估驗單尚未建立
            return contract;
        }

        //代付支出-選商
        public ActionResult ChooseSupplier(string id, string contractid)
        {
            logger.Info("Access To RePayment By EST Form Id =" + id);
            service.getInqueryForm(contractid);
            PLAN_SUP_INQUIRY f = service.formInquiry;
            ViewBag.contractid = contractid;
            ViewBag.formid = id;
            List<RePaymentFunction> lstSupplier = null;
            lstSupplier = service.getSupplierOfContractByPrjId(f.PROJECT_ID);
            ViewBag.status = -10; //估驗單尚未建立
            ViewBag.status = service.getStatusById(id);
            ViewBag.key = lstSupplier.Count;
            logger.Debug("supplier record =" + ViewBag.key + "筆");
            return View(lstSupplier);
        }

        public ActionResult AddRePayment(FormCollection form)
        {
            //取得估驗單編號
            logger.Info("EST Form Id:" + Request["formid"]);
            //取得合約編號
            logger.Info("Contract Id:" + Request["contractid"]);
            //取得使用者勾選品項ID
            logger.Info("item_list:" + Request["chkItem"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            logger.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                logger.Info("item_list return No.:" + lstItemId[i]);
            }
            List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_OTHER_PAYMENT item = new PLAN_OTHER_PAYMENT();
                item.EST_FORM_ID = Request["formid"];
                item.CONTRACT_ID = Request["contractid"];
                item.CONTRACT_ID_FOR_REFUND = lstItemId[j];
                logger.Debug("Item EST form id =" + item.EST_FORM_ID + " ,and supplier for refund =" + item.CONTRACT_ID_FOR_REFUND);
                lstItem.Add(item);
            }
            int k = service.AddRePay(lstItem);
            return Redirect("RePayment?id=" + Request["formid"] + "&contractid=" + Request["contractid"]);
        }
        public String UpdateRePay(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            // 先刪除原先資料
            logger.Info("EST form id =" + form["formid"]);
            logger.Info("Delete PLAN_OTHER_PAYMENT By EST_FORM_ID");
            service.delRePayByESTId(form["formid"]);
            // 再次取得代付支出資料
            string[] lstAmount = form.Get("input_amount").Split(',');
            string[] lstReason = form.Get("input_reason").Split(',');
            string[] lstContract4Refund = form.Get("contractid4refund").Split(',');
            List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
            for (int j = 0; j < lstAmount.Count(); j++)
            {
                PLAN_OTHER_PAYMENT item = new PLAN_OTHER_PAYMENT();
                item.EST_FORM_ID = form["formid"];
                item.CONTRACT_ID = form["contractid"];
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                logger.Info("Other Payment Amount  =" + item.AMOUNT);
                item.REASON = lstReason[j];
                item.CONTRACT_ID_FOR_REFUND = lstContract4Refund[j];
                logger.Debug("Item EST form id =" + item.EST_FORM_ID + "且扣款原因為" + item.REASON + "且代付支出對象為" + item.CONTRACT_ID_FOR_REFUND);
                lstItem.Add(item);
            }
            int i = service.AddRePay(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新代付支出資料成功，EST_FORM_ID =" + form["formid"];
            }
            return msg;
        }
        //代付扣回
        public ActionResult Refund(string id, string contractid)
        {
            logger.Info("Access To Refund By EST Form Id =" + id);
            service.getInqueryForm(contractid);
            PLAN_SUP_INQUIRY f = service.formInquiry;
            ViewBag.projectId = f.PROJECT_ID;
            TND_PROJECT p = service.getProjectById(f.PROJECT_ID);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.contractid = contractid;
            ViewBag.formname = f.FORM_NAME;
            ViewBag.formid = id;
            List<RePaymentFunction> lstOtherPayItem = null;
            List<RePaymentFunction> lstRefundItem = null;
            lstOtherPayItem = service.getRefundById(id);
            lstRefundItem = service.getRefundOfSupplierById(contractid);
            ViewBag.status = -10; //估驗單尚未建立
            ViewBag.status = service.getStatusById(id);
            ViewBag.key = lstOtherPayItem.Count;
            logger.Debug("this repayment record =" + ViewBag.key + "筆");
            ViewData["items"] = JsonConvert.SerializeObject(lstOtherPayItem);
            logger.Debug(ViewData["items"]);
            return View(lstRefundItem);
        }

        //代付支出-選商
        public ActionResult ChooseSupplierOfRefund(string id, string contractid)
        {
            logger.Info("Access To Refund By EST Form Id =" + id);
            service.getInqueryForm(contractid);
            PLAN_SUP_INQUIRY f = service.formInquiry;
            ViewBag.contractid = contractid;
            ViewBag.formid = id;
            List<RePaymentFunction> lstSupplier = null;
            lstSupplier = service.getSupplierOfContractRefundById(contractid);
            ViewBag.status = -10; //估驗單尚未建立
            ViewBag.status = service.getStatusById(id);
            ViewBag.key = lstSupplier.Count;
            logger.Debug("supplier record =" + ViewBag.key + "筆");
            return View(lstSupplier);
        }

        public ActionResult AddRefund(FormCollection form)
        {
            //取得估驗單編號
            logger.Info("EST Form Id:" + Request["formid"]);
            //取得合約編號
            logger.Info("Contract Id:" + Request["contractid"]);
            //取得使用者勾選品項ID
            logger.Info("item_list:" + Request["chkItem"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            logger.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                logger.Info("item_list return No.:" + lstItemId[i]);
            }
            PLAN_OTHER_PAYMENT Item = new PLAN_OTHER_PAYMENT();
            string k = service.AddRefund(Request["formid"], Request["contractid"], lstItemId);
            return Redirect("Refund?id=" + Request["formid"] + "&contractid=" + Request["contractid"]);
        }

        public String UpdateRefund(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            // 先刪除原先資料
            logger.Info("EST form id =" + form["formid"]);
            logger.Info("Delete PLAN_OTHER_PAYMENT By EST_FORM_ID");
            service.delRefundByESTId(form["formid"]);
            // 再次取得代付扣回資料
            string[] lstAmount = form.Get("input_amount").Split(',');
            string[] lstReason = form.Get("input_reason").Split(',');
            string[] lstEstCount = form.Get("est_count").Split(',');
            string[] lstEstIdRefund = form.Get("est_id_refund").Split(',');
            string[] lstContract4Refund = form.Get("contractid4refund").Split(',');
            List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
            for (int j = 0; j < lstAmount.Count(); j++)
            {
                PLAN_OTHER_PAYMENT item = new PLAN_OTHER_PAYMENT();
                item.EST_FORM_ID = form["formid"];
                item.CONTRACT_ID = form["contractid"];
                if (lstEstCount[j].ToString() == "")
                {
                    item.EST_COUNT_REFUND = null;
                }
                else
                {
                    item.EST_COUNT_REFUND = int.Parse(lstEstCount[j]);
                }
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                logger.Info("Other Payment Amount  =" + item.AMOUNT);
                item.REASON = lstReason[j];
                item.CONTRACT_ID_FOR_REFUND = lstContract4Refund[j];
                item.EST_FORM_ID_REFUND = lstEstIdRefund[j];
                logger.Debug("Item EST form id =" + item.EST_FORM_ID + "且扣款原因為" + item.REASON + "且請款對象為" + item.CONTRACT_ID_FOR_REFUND + "且請款單號為" + item.EST_FORM_ID_REFUND);
                lstItem.Add(item);
            }
            int i = service.RefreshRefund(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新代付扣回資料成功，EST_FORM_ID =" + form["formid"];
            }
            return msg;
        }

        //取得未核准的估驗單
        public string getEstNoOfUnApproval()
        {
            logger.Debug("get EST No of un Approval By contractid:" + Request["contractid"]);
            string contractid = Request["contractid"];
            string UnApproval = null;
            UnApproval = service.getEstNoByContractId(contractid);
            return UnApproval;
        }

        //工地費用功能
        #region 工地費用
        public ActionResult Expense(string id)
        {
            logger.Info("Access to Expense Page !!");
            ViewBag.projectid = id;
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            List<FIN_SUBJECT> Subject = null;
            Subject = service.getSubjectOfExpense4Site();
            ViewData["items"] = JsonConvert.SerializeObject(Subject);
            return View();
        }

        public ActionResult SearchSubject()
        {
            //取得使用者勾選品項ID
            logger.Info("item_list:" + Request["subject"]);
            string[] lstItemId = Request["subject"].ToString().Split(',');
            logger.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                logger.Info("item_list return No.:" + lstItemId[i]);
            }
            ViewBag.projectid = Request["projectid"];
            TND_PROJECT p = service.getProjectById(Request["projectid"]);
            ViewBag.projectName = p.PROJECT_NAME;
            List<FIN_SUBJECT> SubjectChecked = null;
            SubjectChecked = service.getSubjectByChkItem(lstItemId);
            List<FIN_SUBJECT> Subject = null;
            Subject = service.getSubjectOfExpense4Site();
            ViewData["items"] = JsonConvert.SerializeObject(Subject);
            return View("Expense", SubjectChecked);
        }

        [HttpPost]
        public ActionResult AddExpense(FIN_EXPENSE_FORM ef)
        {
            string[] lstSubject = Request["subject"].Split(',');
            string[] lstAmount = Request["expense_amount"].Split(',');
            string[] lstRemark = Request["item_remark"].Split(',');
            //建立工地費用單號
            logger.Info("create new Plan Expense Form");
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            ef.PROJECT_ID = Request["projectid"];
            ef.PAYMENT_DATE = Convert.ToDateTime(Request["paymentdate"]);
            ef.OCCURRED_YEAR = int.Parse(Request["paymentdate"].Substring(0, 4));
            ef.OCCURRED_MONTH = int.Parse(Request["paymentdate"].Substring(5, 2));
            ef.CREATE_DATE = DateTime.Now;
            ef.CREATE_ID = uInfo.USER_ID;
            ef.REMARK = Request["remark"];
            ef.STATUS = 10;
            string fid = service.newExpenseForm(ef);
            //建立工地費用單明細
            List<FIN_EXPENSE_ITEM> lstItem = new List<FIN_EXPENSE_ITEM>();
            for (int j = 0; j < lstSubject.Count(); j++)
            {
                FIN_EXPENSE_ITEM item = new FIN_EXPENSE_ITEM();
                item.FIN_SUBJECT_ID = lstSubject[j];
                item.ITEM_REMARK = lstRemark[j];
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                logger.Info("Plan Expense Subject =" + item.FIN_SUBJECT_ID + "， and Amount = " + item.AMOUNT);
                item.EXP_FORM_ID = fid;
                logger.Debug("Item EX form id =" + item.EXP_FORM_ID);
                lstItem.Add(item);
            }
            int i = service.AddExpenseItems(lstItem);
            //建立申請單參考流程
            Flow4SiteExpense flowService = new Flow4SiteExpense();
            logger.Debug("Item Count =" + i);
            flowService.iniRequest(u, fid);
            return RedirectToAction("SingleEXPForm", "CashFlow", new { id = fid });
        }

        public ActionResult SiteBudget(string id)
        {
            logger.Info("Access to Site Budget Page, And Project Id = " + id);
            ViewBag.projectid = id;
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            //取得工地費用預算資料
            var priId = service.getSiteBudgetById(id);
            ViewBag.budgetdata = priId;
            List<ExpenseBudgetSummary> FirstYearBudget = null;
            List<ExpenseBudgetSummary> SecondYearBudget = null;
            List<ExpenseBudgetSummary> ThirdYearBudget = null;
            List<ExpenseBudgetSummary> FourthYearBudget = null;
            List<ExpenseBudgetSummary> FifthYearBudget = null;
            ExpenseBudgetSummary Amt = null;
            if (null != priId && priId != "")
            {
                //取得已寫入之工地費用預算資料
                FirstYearBudget = service.getBudget4ProjectBySeq(id, null, "1");
                SecondYearBudget = service.getBudget4ProjectBySeq(id, null, "2");
                ThirdYearBudget = service.getBudget4ProjectBySeq(id, null, "3");
                FourthYearBudget = service.getBudget4ProjectBySeq(id, null, "4");
                FifthYearBudget = service.getBudget4ProjectBySeq(id, null, "5");
                //工地(專案)該年度預算
                ViewBag.FirstYear = service.getSiteBudgetByYearSeq(id, "1");
                ViewBag.SecondYear = service.getSiteBudgetByYearSeq(id, "2");
                ViewBag.ThirdYear = service.getSiteBudgetByYearSeq(id, "3");
                ViewBag.FourthYear = service.getSiteBudgetByYearSeq(id, "4");
                ViewBag.FifthYear = service.getSiteBudgetByYearSeq(id, "5");
                Amt = service.getSiteBudgetAmountById(id, null);
                TempData["TotalAmt"] = Amt.TOTAL_BUDGET;
                SiteBudgetModels viewModel = new SiteBudgetModels();
                viewModel.firstYear = FirstYearBudget;
                viewModel.secondYear = SecondYearBudget;
                viewModel.thirdYear = ThirdYearBudget;
                viewModel.fourthYear = FourthYearBudget;
                viewModel.fifthYear = FifthYearBudget;
                return View(viewModel);
            }
            return View();
        }

        //更新工地費用預算
        #region 第1年度
        public String UpdateSiteBudgetOfFirstYear(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            // 先刪除原先資料
            logger.Info("Site Expense Budget's Project Id =" + form["projectId"]);
            logger.Info("Delete PLAN_SITE_BUDGET By PROJECT_ID and YEAR_SEQUENCE");
            var year = 1.ToString();
            service.delSiteBudgetByProject(form["projectId"], year);
            string msg = "";
            string[] lstsubjctid = form.Get("subjctid").Split(',');
            string[] lst1 = form.Get("janAmt1").Split(',');
            string[] lst2 = form.Get("febAmt1").Split(',');
            string[] lst3 = form.Get("marAmt1").Split(',');
            string[] lst4 = form.Get("aprAmt1").Split(',');
            string[] lst5 = form.Get("mayAmt1").Split(',');
            string[] lst6 = form.Get("junAmt1").Split(',');
            string[] lst7 = form.Get("julAmt1").Split(',');
            string[] lst8 = form.Get("augAmt1").Split(',');
            string[] lst9 = form.Get("sepAmt1").Split(',');
            string[] lst10 = form.Get("octAmt1").Split(',');
            string[] lst11 = form.Get("novAmt1").Split(',');
            string[] lst12 = form.Get("decAmt1").Split(',');
            List<string[]> Atm = new List<string[]>();
            Atm.Add(lst1);
            Atm.Add(lst2);
            Atm.Add(lst3);
            Atm.Add(lst4);
            Atm.Add(lst5);
            Atm.Add(lst6);
            Atm.Add(lst7);
            Atm.Add(lst8);
            Atm.Add(lst9);
            Atm.Add(lst10);
            Atm.Add(lst11);
            Atm.Add(lst12);
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            List<PLAN_SITE_BUDGET> lst = new List<PLAN_SITE_BUDGET>();
            for (int j = 0; j < lstsubjctid.Count(); j++)
            {
                List<PLAN_SITE_BUDGET> lstItem = new List<PLAN_SITE_BUDGET>();
                for (int i = 0; i < 12; i++)
                {
                    PLAN_SITE_BUDGET item = new PLAN_SITE_BUDGET();
                    item.BUDGET_YEAR = int.Parse(form["firstYear"]);
                    item.PROJECT_ID = form["projectId"];
                    item.YEAR_SEQUENCE = year;
                    item.SUBJECT_ID = lstsubjctid[j];
                    item.MODIFY_ID = u.USER_ID;
                    if (Atm[i][j].ToString() == "" && null != Atm[i][j].ToString())
                    {
                        item.AMOUNT = null;
                        item.BUDGET_MONTH = i + 1;
                    }
                    else
                    {
                        item.AMOUNT = decimal.Parse(Atm[i][j]);
                        item.BUDGET_MONTH = i + 1;
                    }
                    item.MODIFY_DATE = DateTime.Now;
                    logger.Info("費用項目代碼 =" + item.SUBJECT_ID + "，and Budget Month = " + item.BUDGET_MONTH + "，and Atm = " + Atm[i][j]);
                    lstItem.Add(item);
                }
                lst.AddRange(lstItem);
            }
            int k = service.refreshSiteBudget(lst);
            if (k == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新工地預算費用成功，預算年度為 " + form["firstYear"];
            }

            logger.Info("Request:BUDGET_YEAR =" + form["firstYear"]);
            return msg;
        }
        #endregion
        #region 第2年度
        public String UpdateSiteBudgetOfSecondYear(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            // 先刪除原先資料
            logger.Info("Site Expense Budget's Project Id =" + form["projectId"]);
            logger.Info("Delete PLAN_SITE_BUDGET By PROJECT_ID and YEAR_SEQUENCE");
            var year = 2.ToString();
            service.delSiteBudgetByProject(form["projectId"], year);
            string msg = "";
            string[] lstsubjctid = form.Get("subjctid").Split(',');
            string[] lst1 = form.Get("janAmt2").Split(',');
            string[] lst2 = form.Get("febAmt2").Split(',');
            string[] lst3 = form.Get("marAmt2").Split(',');
            string[] lst4 = form.Get("aprAmt2").Split(',');
            string[] lst5 = form.Get("mayAmt2").Split(',');
            string[] lst6 = form.Get("junAmt2").Split(',');
            string[] lst7 = form.Get("julAmt2").Split(',');
            string[] lst8 = form.Get("augAmt2").Split(',');
            string[] lst9 = form.Get("sepAmt2").Split(',');
            string[] lst10 = form.Get("octAmt2").Split(',');
            string[] lst11 = form.Get("novAmt2").Split(',');
            string[] lst12 = form.Get("decAmt2").Split(',');
            List<string[]> Atm = new List<string[]>();
            Atm.Add(lst1);
            Atm.Add(lst2);
            Atm.Add(lst3);
            Atm.Add(lst4);
            Atm.Add(lst5);
            Atm.Add(lst6);
            Atm.Add(lst7);
            Atm.Add(lst8);
            Atm.Add(lst9);
            Atm.Add(lst10);
            Atm.Add(lst11);
            Atm.Add(lst12);
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            List<PLAN_SITE_BUDGET> lst = new List<PLAN_SITE_BUDGET>();
            for (int j = 0; j < lstsubjctid.Count(); j++)
            {
                List<PLAN_SITE_BUDGET> lstItem = new List<PLAN_SITE_BUDGET>();
                for (int i = 0; i < 12; i++)
                {
                    PLAN_SITE_BUDGET item = new PLAN_SITE_BUDGET();
                    item.BUDGET_YEAR = int.Parse(form["secondYear"]);
                    item.PROJECT_ID = form["projectId"];
                    item.YEAR_SEQUENCE = year;
                    item.SUBJECT_ID = lstsubjctid[j];
                    item.MODIFY_ID = u.USER_ID;
                    if (Atm[i][j].ToString() == "" && null != Atm[i][j].ToString())
                    {
                        item.AMOUNT = null;
                        item.BUDGET_MONTH = i + 1;
                    }
                    else
                    {
                        item.AMOUNT = decimal.Parse(Atm[i][j]);
                        item.BUDGET_MONTH = i + 1;
                    }
                    item.MODIFY_DATE = DateTime.Now;
                    logger.Info("費用項目代碼 =" + item.SUBJECT_ID + "，and Budget Month = " + item.BUDGET_MONTH + "，and Atm = " + Atm[i][j]);
                    lstItem.Add(item);
                }
                lst.AddRange(lstItem);
            }
            int k = service.refreshSiteBudget(lst);
            if (k == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新工地預算費用成功，預算年度為 " + form["secondYear"];
            }

            logger.Info("Request:BUDGET_YEAR =" + form["secondYear"]);
            return msg;
        }
        #endregion
        #region 第3年度
        public String UpdateSiteBudgetOfThirdYear(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            // 先刪除原先資料
            logger.Info("Site Expense Budget's Project Id =" + form["projectId"]);
            logger.Info("Delete PLAN_SITE_BUDGET By PROJECT_ID and YEAR_SEQUENCE");
            var year = 3.ToString();
            service.delSiteBudgetByProject(form["projectId"], year);
            string msg = "";
            string[] lstsubjctid = form.Get("subjctid").Split(',');
            string[] lst1 = form.Get("janAmt3").Split(',');
            string[] lst2 = form.Get("febAmt3").Split(',');
            string[] lst3 = form.Get("marAmt3").Split(',');
            string[] lst4 = form.Get("aprAmt3").Split(',');
            string[] lst5 = form.Get("mayAmt3").Split(',');
            string[] lst6 = form.Get("junAmt3").Split(',');
            string[] lst7 = form.Get("julAmt3").Split(',');
            string[] lst8 = form.Get("augAmt3").Split(',');
            string[] lst9 = form.Get("sepAmt3").Split(',');
            string[] lst10 = form.Get("octAmt3").Split(',');
            string[] lst11 = form.Get("novAmt3").Split(',');
            string[] lst12 = form.Get("decAmt3").Split(',');
            List<string[]> Atm = new List<string[]>();
            Atm.Add(lst1);
            Atm.Add(lst2);
            Atm.Add(lst3);
            Atm.Add(lst4);
            Atm.Add(lst5);
            Atm.Add(lst6);
            Atm.Add(lst7);
            Atm.Add(lst8);
            Atm.Add(lst9);
            Atm.Add(lst10);
            Atm.Add(lst11);
            Atm.Add(lst12);
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            List<PLAN_SITE_BUDGET> lst = new List<PLAN_SITE_BUDGET>();
            for (int j = 0; j < lstsubjctid.Count(); j++)
            {
                List<PLAN_SITE_BUDGET> lstItem = new List<PLAN_SITE_BUDGET>();
                for (int i = 0; i < 12; i++)
                {
                    PLAN_SITE_BUDGET item = new PLAN_SITE_BUDGET();
                    item.BUDGET_YEAR = int.Parse(form["thirdYear"]);
                    item.PROJECT_ID = form["projectId"];
                    item.YEAR_SEQUENCE = year;
                    item.SUBJECT_ID = lstsubjctid[j];
                    item.MODIFY_ID = u.USER_ID;
                    if (Atm[i][j].ToString() == "" && null != Atm[i][j].ToString())
                    {
                        item.AMOUNT = null;
                        item.BUDGET_MONTH = i + 1;
                    }
                    else
                    {
                        item.AMOUNT = decimal.Parse(Atm[i][j]);
                        item.BUDGET_MONTH = i + 1;
                    }
                    item.MODIFY_DATE = DateTime.Now;
                    logger.Info("費用項目代碼 =" + item.SUBJECT_ID + "，and Budget Month = " + item.BUDGET_MONTH + "，and Atm = " + Atm[i][j]);
                    lstItem.Add(item);
                }
                lst.AddRange(lstItem);
            }
            int k = service.refreshSiteBudget(lst);
            if (k == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新工地預算費用成功，預算年度為 " + form["thirdYear"];
            }

            logger.Info("Request:BUDGET_YEAR =" + form["thirdYear"]);
            return msg;
        }
        #endregion
        #region 第4年度
        public String UpdateSiteBudgetOfFourthYear(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            // 先刪除原先資料
            logger.Info("Site Expense Budget's Project Id =" + form["projectId"]);
            logger.Info("Delete PLAN_SITE_BUDGET By PROJECT_ID and YEAR_SEQUENCE");
            var year = 4.ToString();
            service.delSiteBudgetByProject(form["projectId"], year);
            string msg = "";
            string[] lstsubjctid = form.Get("subjctid").Split(',');
            string[] lst1 = form.Get("janAmt4").Split(',');
            string[] lst2 = form.Get("febAmt4").Split(',');
            string[] lst3 = form.Get("marAmt4").Split(',');
            string[] lst4 = form.Get("aprAmt4").Split(',');
            string[] lst5 = form.Get("mayAmt4").Split(',');
            string[] lst6 = form.Get("junAmt4").Split(',');
            string[] lst7 = form.Get("julAmt4").Split(',');
            string[] lst8 = form.Get("augAmt4").Split(',');
            string[] lst9 = form.Get("sepAmt4").Split(',');
            string[] lst10 = form.Get("octAmt4").Split(',');
            string[] lst11 = form.Get("novAmt4").Split(',');
            string[] lst12 = form.Get("decAmt4").Split(',');
            List<string[]> Atm = new List<string[]>();
            Atm.Add(lst1);
            Atm.Add(lst2);
            Atm.Add(lst3);
            Atm.Add(lst4);
            Atm.Add(lst5);
            Atm.Add(lst6);
            Atm.Add(lst7);
            Atm.Add(lst8);
            Atm.Add(lst9);
            Atm.Add(lst10);
            Atm.Add(lst11);
            Atm.Add(lst12);
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            List<PLAN_SITE_BUDGET> lst = new List<PLAN_SITE_BUDGET>();
            for (int j = 0; j < lstsubjctid.Count(); j++)
            {
                List<PLAN_SITE_BUDGET> lstItem = new List<PLAN_SITE_BUDGET>();
                for (int i = 0; i < 12; i++)
                {
                    PLAN_SITE_BUDGET item = new PLAN_SITE_BUDGET();
                    item.BUDGET_YEAR = int.Parse(form["fourthYear"]);
                    item.PROJECT_ID = form["projectId"];
                    item.YEAR_SEQUENCE = year;
                    item.SUBJECT_ID = lstsubjctid[j];
                    item.MODIFY_ID = u.USER_ID;
                    if (Atm[i][j].ToString() == "" && null != Atm[i][j].ToString())
                    {
                        item.AMOUNT = null;
                        item.BUDGET_MONTH = i + 1;
                    }
                    else
                    {
                        item.AMOUNT = decimal.Parse(Atm[i][j]);
                        item.BUDGET_MONTH = i + 1;
                    }
                    item.MODIFY_DATE = DateTime.Now;
                    logger.Info("費用項目代碼 =" + item.SUBJECT_ID + "，and Budget Month = " + item.BUDGET_MONTH + "，and Atm = " + Atm[i][j]);
                    lstItem.Add(item);
                }
                lst.AddRange(lstItem);
            }
            int k = service.refreshSiteBudget(lst);
            if (k == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新工地預算費用成功，預算年度為 " + form["fourthYear"];
            }

            logger.Info("Request:BUDGET_YEAR =" + form["fourthYear"]);
            return msg;
        }
        #endregion
        #region 第5年度
        public String UpdateSiteBudgetOfFifthYear(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            // 先刪除原先資料
            logger.Info("Site Expense Budget's Project Id =" + form["projectId"]);
            logger.Info("Delete PLAN_SITE_BUDGET By PROJECT_ID and YEAR_SEQUENCE");
            var year = 4.ToString();
            service.delSiteBudgetByProject(form["projectId"], year);
            string msg = "";
            string[] lstsubjctid = form.Get("subjctid").Split(',');
            string[] lst1 = form.Get("janAmt5").Split(',');
            string[] lst2 = form.Get("febAmt5").Split(',');
            string[] lst3 = form.Get("marAmt5").Split(',');
            string[] lst4 = form.Get("aprAmt5").Split(',');
            string[] lst5 = form.Get("mayAmt5").Split(',');
            string[] lst6 = form.Get("junAmt5").Split(',');
            string[] lst7 = form.Get("julAmt5").Split(',');
            string[] lst8 = form.Get("augAmt5").Split(',');
            string[] lst9 = form.Get("sepAmt5").Split(',');
            string[] lst10 = form.Get("octAmt5").Split(',');
            string[] lst11 = form.Get("novAmt5").Split(',');
            string[] lst12 = form.Get("decAmt5").Split(',');
            List<string[]> Atm = new List<string[]>();
            Atm.Add(lst1);
            Atm.Add(lst2);
            Atm.Add(lst3);
            Atm.Add(lst4);
            Atm.Add(lst5);
            Atm.Add(lst6);
            Atm.Add(lst7);
            Atm.Add(lst8);
            Atm.Add(lst9);
            Atm.Add(lst10);
            Atm.Add(lst11);
            Atm.Add(lst12);
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            List<PLAN_SITE_BUDGET> lst = new List<PLAN_SITE_BUDGET>();
            for (int j = 0; j < lstsubjctid.Count(); j++)
            {
                List<PLAN_SITE_BUDGET> lstItem = new List<PLAN_SITE_BUDGET>();
                for (int i = 0; i < 12; i++)
                {
                    PLAN_SITE_BUDGET item = new PLAN_SITE_BUDGET();
                    item.BUDGET_YEAR = int.Parse(form["fourthYear"]);
                    item.PROJECT_ID = form["projectId"];
                    item.YEAR_SEQUENCE = year;
                    item.SUBJECT_ID = lstsubjctid[j];
                    item.MODIFY_ID = u.USER_ID;
                    if (Atm[i][j].ToString() == "" && null != Atm[i][j].ToString())
                    {
                        item.AMOUNT = null;
                        item.BUDGET_MONTH = i + 1;
                    }
                    else
                    {
                        item.AMOUNT = decimal.Parse(Atm[i][j]);
                        item.BUDGET_MONTH = i + 1;
                    }
                    item.MODIFY_DATE = DateTime.Now;
                    logger.Info("費用項目代碼 =" + item.SUBJECT_ID + "，and Budget Month = " + item.BUDGET_MONTH + "，and Atm = " + Atm[i][j]);
                    lstItem.Add(item);
                }
                lst.AddRange(lstItem);
            }
            int k = service.refreshSiteBudget(lst);
            if (k == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新工地預算費用成功，預算年度為 " + form["fifthYear"];
            }

            logger.Info("Request:BUDGET_YEAR =" + form["fifthYear"]);
            return msg;
        }
        #endregion

        //工地費用下載與上傳操作頁面
        public ActionResult SiteBudgetOperation(string id)
        {
            logger.Info("Access to Site Budget Page, And Project Id = " + id);
            ViewBag.projectid = id;
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            return View();
        }
        /// <summary>
        /// 下載工地費用預算填寫表
        /// </summary>
        public void downLoadSiteBudgetForm()
        {
            string projectid = Request["projectid"];
            PlanService ps = new PlanService();
            ps.getProjectId(projectid);
            if (null != ps.budgetTable)
            {
                SiteBudgetFormToExcel poi = new SiteBudgetFormToExcel();
                //檔案位置
                string fileLocation = poi.exportExcel(ps.budgetTable);
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

        //上傳工地費用預算
        [HttpPost]
        public ActionResult uploadSiteBudgetTable(HttpPostedFileBase fileBudget1, HttpPostedFileBase fileBudget2, HttpPostedFileBase fileBudget3,
            HttpPostedFileBase fileBudget4, HttpPostedFileBase fileBudget5)
        {
            string projectid = Request["projectid"];
            string message = "";
            logger.Info("Upload plan Budget of Site Expenses for projectid=" + projectid);
            try
            {
                //檔案變數名稱需要與前端畫面對應
                #region 預算第1年度
                //工地費用:預算第1年度
                if (null != fileBudget1 && fileBudget1.ContentLength != 0)
                {
                    //2.解析Excel
                    logger.Info("Parser Excel data:" + fileBudget1.FileName);
                    //2.1 設定Excel 檔案名稱
                    var fileName = Path.GetFileName(fileBudget1.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    logger.Info("save excel file:" + path);
                    fileBudget1.SaveAs(path);
                    //2.2 開啟Excel 檔案
                    logger.Info("Parser Excel File Begin:" + fileBudget1.FileName);
                    SiteBudgetFormToExcel budgetservice = new SiteBudgetFormToExcel();
                    budgetservice.InitializeWorkbook(path);
                    //解析預算數量
                    List<PLAN_SITE_BUDGET> lstSiteBudget = budgetservice.ConvertDataForSiteBudget1(projectid);
                    //2.3 記錄錯誤訊息
                    message = budgetservice.errorMessage;
                    //2.4
                    var year = 1.ToString();
                    logger.Info("Delete PALN_SITE_BUDGET By Project");
                    service.delSiteBudgetByProject(projectid, year);
                    message = message + "<br/>舊有資料刪除成功 !!";
                    //2.5 
                    logger.Info("Add First Year PALN_SITE_BUDGET to DB");
                    service.refreshSiteBudget(lstSiteBudget);
                    message = message + "<br/>資料匯入完成 !!";
                }
                #endregion
                #region 預算第2年度
                //工地費用:預算第2年度

                if (null != fileBudget2 && fileBudget2.ContentLength != 0)
                {
                    //2.解析Excel

                    logger.Info("Parser Excel data:" + fileBudget2.FileName);
                    //2.1 設定Excel 檔案名稱
                    var fileName = Path.GetFileName(fileBudget2.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    logger.Info("save excel file:" + path);
                    fileBudget2.SaveAs(path);
                    //2.2 開啟Excel 檔案
                    logger.Info("Parser Excel File Begin:" + fileBudget2.FileName);
                    SiteBudgetFormToExcel budgetservice = new SiteBudgetFormToExcel();
                    budgetservice.InitializeWorkbook(path);
                    //解析預算數量
                    List<PLAN_SITE_BUDGET> _2ndSiteBudget = budgetservice.ConvertDataForSiteBudget2(projectid);
                    //2.3 記錄錯誤訊息
                    message = budgetservice.errorMessage;
                    //2.4
                    var year = 2.ToString();
                    logger.Info("Delete PALN_SITE_BUDGET By Project");
                    service.delSiteBudgetByProject(projectid, year);
                    message = message + "<br/>舊有資料刪除成功 !!";
                    //2.5 
                    logger.Info("Add Second Year PALN_SITE_BUDGET to DB");
                    service.refreshSiteBudget(_2ndSiteBudget);
                    message = message + "<br/>資料匯入完成 !!";
                }
                #endregion
                #region 預算第3年度
                //工地費用:預算第3年度

                if (null != fileBudget3 && fileBudget3.ContentLength != 0)
                {
                    //2.解析Excel
                    logger.Info("Parser Excel data:" + fileBudget3.FileName);
                    //2.1 設定Excel 檔案名稱
                    var fileName = Path.GetFileName(fileBudget3.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    logger.Info("save excel file:" + path);
                    fileBudget3.SaveAs(path);
                    //2.2 開啟Excel 檔案
                    logger.Info("Parser Excel File Begin:" + fileBudget3.FileName);
                    SiteBudgetFormToExcel budgetservice = new SiteBudgetFormToExcel();
                    budgetservice.InitializeWorkbook(path);
                    //解析預算數量
                    List<PLAN_SITE_BUDGET> _3rdSiteBudget = budgetservice.ConvertDataForSiteBudget3(projectid);
                    //2.3 記錄錯誤訊息
                    message = budgetservice.errorMessage;
                    //2.4
                    var year = 3.ToString();
                    logger.Info("Delete PALN_SITE_BUDGET By Project");
                    service.delSiteBudgetByProject(projectid, year);
                    message = message + "<br/>舊有資料刪除成功 !!";
                    //2.5 
                    logger.Info("Add Second Year PALN_SITE_BUDGET to DB");
                    service.refreshSiteBudget(_3rdSiteBudget);
                    message = message + "<br/>資料匯入完成 !!";
                }
                #endregion
                #region 預算第4年度
                //工地費用:預算第4年度

                if (null != fileBudget4 && fileBudget4.ContentLength != 0)
                {
                    //2.解析Excel
                    logger.Info("Parser Excel data:" + fileBudget4.FileName);
                    //2.1 設定Excel 檔案名稱
                    var fileName = Path.GetFileName(fileBudget4.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    logger.Info("save excel file:" + path);
                    fileBudget4.SaveAs(path);
                    //2.2 開啟Excel 檔案
                    logger.Info("Parser Excel File Begin:" + fileBudget4.FileName);
                    SiteBudgetFormToExcel budgetservice = new SiteBudgetFormToExcel();
                    budgetservice.InitializeWorkbook(path);
                    //解析預算數量
                    List<PLAN_SITE_BUDGET> _4thSiteBudget = budgetservice.ConvertDataForSiteBudget3(projectid);
                    //2.3 記錄錯誤訊息
                    message = budgetservice.errorMessage;
                    //2.4
                    var year = 4.ToString();
                    logger.Info("Delete PALN_SITE_BUDGET By Project");
                    service.delSiteBudgetByProject(projectid, year);
                    message = message + "<br/>舊有資料刪除成功 !!";
                    //2.5 
                    logger.Info("Add Second Year PALN_SITE_BUDGET to DB");
                    service.refreshSiteBudget(_4thSiteBudget);
                    message = message + "<br/>資料匯入完成 !!";
                }
                #endregion
                #region 預算第5年度
                //工地費用:預算第5年度

                if (null != fileBudget5 && fileBudget5.ContentLength != 0)
                {
                    //2.解析Excel
                    logger.Info("Parser Excel data:" + fileBudget5.FileName);
                    //2.1 設定Excel 檔案名稱
                    var fileName = Path.GetFileName(fileBudget5.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    logger.Info("save excel file:" + path);
                    fileBudget5.SaveAs(path);
                    //2.2 開啟Excel 檔案
                    logger.Info("Parser Excel File Begin:" + fileBudget5.FileName);
                    SiteBudgetFormToExcel budgetservice = new SiteBudgetFormToExcel();
                    budgetservice.InitializeWorkbook(path);
                    //解析預算數量
                    List<PLAN_SITE_BUDGET> _5thSiteBudget = budgetservice.ConvertDataForSiteBudget5(projectid);
                    //2.3 記錄錯誤訊息
                    message = budgetservice.errorMessage;
                    //2.4
                    var year = 5.ToString();
                    logger.Info("Delete PLAN_SITE_BUDGET By Project");
                    service.delSiteBudgetByProject(projectid, year);
                    message = message + "<br/>舊有資料刪除成功 !!";
                    //2.5 
                    logger.Info("Add Second Year PALN_SITE_BUDGET to DB");
                    service.refreshSiteBudget(_5thSiteBudget);
                    message = message + "<br/>資料匯入完成 !!";
                }
                #endregion
            }
            catch (Exception ex)
            {
                logger.Error(ex.StackTrace);
                message = ex.Message;
            }
            TempData["result"] = message;
            return RedirectToAction("SiteBudget/" + projectid);
        }
        #endregion
        // 工地預算查詢畫面
        public ActionResult SiteExpSummary(string id)
        {
            logger.Info("Access to Site Expense and Budget Summary Page，Project Id =" + id);
            ViewBag.projectid = id;
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            //取得工地每年預算資料
            List<ExpenseBudgetSummary> lstSiteBudget = service.getSiteBudgetPerYear(id);
            ExpenseBudgetModel viewModel = new ExpenseBudgetModel();
            //取得工地總預算、費用累積
            ExpenseBudgetSummary BudgeTotalAmt = service.getSiteBudgetAmountById(id, null);
            ExpenseBudgetSummary ExpenseTotalAmt = service.getTotalSiteExpAmountById(id, null);
            TempData["TotalBudget"] = BudgeTotalAmt.TOTAL_BUDGET; //getTotalSiteBudgetAmount
            TempData["TotalExpense"] = ExpenseTotalAmt.CUM_YEAR_AMOUNT;
            viewModel.SiteBudgetPerYear = lstSiteBudget;
            return View(viewModel);
        }
        // 工地預算查詢畫面
        public ActionResult SearchSiteExpSummary()
        {
            logger.Info("Access to Site Expense and Budget Summary Page，Project Id =" + Request["projectid"]);
            string projectId = Request["projectid"];
            string targetYear = Request["budgetYear"];

            //取得預算數
            List<ExpenseBudgetSummary> siteBudgetSummary = null;
            //取得發生數
            List<ExpenseBudgetSummary> sitExpenseSummary = null;
            ExpenseBudgetModel viewModel = new ExpenseBudgetModel();
            if (targetYear != null && targetYear != "")
            {
                siteBudgetSummary = service.getBudget4ProjectBySeq(Request["projectid"], targetYear, null);
                sitExpenseSummary = service.getSiteExpenseSummaryByYear(Request["projectid"], targetYear);
                //取得每年工地總預算、費用累積
                ExpenseBudgetSummary BudgeTotalAmt = service.getSiteBudgetAmountById(projectId, targetYear);
                ExpenseBudgetSummary ExpenseTotalAmt = service.getTotalSiteExpAmountById(projectId, targetYear);

                TempData["TotalBudgetPerYear"] = BudgeTotalAmt.TOTAL_BUDGET; //getTotalSiteBudgetAmount
                TempData["TotalExpensePerYear"] = ExpenseTotalAmt.CUM_YEAR_AMOUNT;
                viewModel.BudgetSummary = siteBudgetSummary;
                viewModel.ExpenseSummary = sitExpenseSummary;
                ViewBag.BudgetYear = targetYear;
                return PartialView("_SiteExpBudgetList", viewModel);
            }
            return PartialView("_SiteExpBudgetList");
        }

        /// <summary>
        ///  工地費用預算執行彙整表
        /// </summary>
        public void downLoadSiteExpenseSummary()
        {
            string projectid = Request["projectid"];

            string targetYear = Request["targetYear"];
            SiteExpSummaryToExcel poi = new SiteExpSummaryToExcel();
            //檔案位置
            TnderProjectService t = new TnderProjectService();
            TND_PROJECT p = t.getProjectById(projectid);
            string fileLocation = poi.exportExcel(p.PROJECT_ID, p.PROJECT_NAME);
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

        //業主估驗計價
        public ActionResult Valuation4Owner(string id)
        {
            logger.Info("Access to Valuation For Owner Page，Project Id =" + id);
            ViewBag.projectid = id;
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            //取得付款條件
            PaymentTermsFunction payment = service.getPaymentTerm(id, id);
            if (null == payment)
            {
                payment = new PaymentTermsFunction();
            }
            if (payment.PAYMENT_RETENTION_RATIO != null)
            {
                ViewBag.retention = (null == payment.PAYMENT_RETENTION_RATIO ? 0 : payment.PAYMENT_RETENTION_RATIO);
            }
            else if (payment.USANCE_RETENTION_RATIO != null)
            {
                ViewBag.retention = (null == payment.USANCE_RETENTION_RATIO ? 0 : payment.USANCE_RETENTION_RATIO);
            }
            else { ViewBag.retention = 0; }
            if (payment.PAYMENT_ADVANCE_RATIO != null)
            {
                ViewBag.advance = (null == payment.PAYMENT_ADVANCE_RATIO ? 0 : payment.PAYMENT_ADVANCE_RATIO);
            }
            else if (payment.USANCE_ADVANCE_RATIO != null)
            {
                ViewBag.advance = (null == payment.USANCE_ADVANCE_RATIO ? 0 : payment.USANCE_ADVANCE_RATIO);
            }
            else { ViewBag.advance = 0; }
            service.getAllBankLoan(id);
            ViewData.Add("loans", service.cashFlowModel.finLoan);
            if (service.cashFlowModel.finLoan != null)
            {
                SelectList loans = new SelectList(service.cashFlowModel.finLoan, "BL_ID", "ACCOUNT_NAME");

                ViewBag.loans = loans;
                //將資料存入TempData 減少不斷讀取資料庫
                TempData.Remove("loans");
                TempData.Add("loans", service.cashFlowModel.finLoan);
            }
            RevenueFromOwner va = service.getVACount4OwnerById(id);
            ViewBag.VACount = va.isVA;
            if (va.isVA > 1)
            {
                List<RevenueFromOwner> valuation = null;
                RevenueFromOwner summary = service.getVASummaryAtmById(id);
                ViewBag.contractAtm = (null == summary.contractAtm ? 0 : summary.contractAtm);
                ViewBag.advancePaymentBalance = (null == summary.advancePaymentBalance ? 0 : summary.advancePaymentBalance);
                ViewBag.totalTax = (null == summary.taxAmt ? 0 : summary.taxAmt);
                ViewBag.totalRetention = (null == summary.RETENTION_PAYMENT ? 0 : summary.RETENTION_PAYMENT);
                ViewBag.VAAtm = (null == summary.VALUATION_AMOUNT ? 0 : summary.VALUATION_AMOUNT);
                ViewBag.AR = (null == summary.AR ? 0 : summary.AR);
                ViewBag.ARUnPaid = (null == summary.AR ? 0 : summary.AR) - (null == summary.AR_PAID ? 0 : summary.AR_PAID);
                ViewBag.VABalance = (null == summary.contractAtm ? 0 : summary.contractAtm) - (null == summary.VALUATION_AMOUNT ? 0 : summary.VALUATION_AMOUNT);
                valuation = service.getVADetailById(id);

                return View(valuation);
            }
            return View();
        }
        /// <summary>
        /// 取得業主計價次數
        /// </summary>
        /// <param name="projectid"></param>
        /// <returns></returns>
        public string getVAItem(string projectid)
        {
            logger.Info("get VA item by project id=" + projectid);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(service.getVACount4OwnerById(projectid));
            logger.Info("VA item's Info=" + itemJson);
            return itemJson;
        }
        /// <summary>
        /// 取得業主計價資料
        /// </summary>
        /// <param name="formid"></param>
        /// <returns></returns>
        public string getVADetail(string formid)
        {
            logger.Info("get VA detail by form id=" + formid);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(service.getVADetailByVAId(formid));
            logger.Info("VA detail's Info=" + itemJson);
            return itemJson;
        }
        //新增業主計價資料
        public ActionResult VAItem(String id, String formid)
        {
            logger.Info("get va item for update:project_id=" + id);
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.supplier = p.OWNER_NAME;
            ViewBag.projectId = id;
            List<PLAN_INVOICE> lstInvoice = null;
            lstInvoice = service.getInvoiceById(formid);
            List<CreditNote> lstNote = service.getCreditNoteById(id, formid);
            ViewBag.InvoicePieces = service.getInvoicePiecesById(formid);
            ViewBag.NotePieces = lstNote.Count;
            ViewData["items"] = JsonConvert.SerializeObject(lstInvoice);

            if (id != formid)
            {
                RevenueFromOwner v = service.getVADetailByVAId(formid);
                ViewBag.vaAmt = v.VALUATION_AMOUNT;
                ViewBag.advance = v.ADVANCE_PAYMENT;
                ViewBag.retention = v.RETENTION_PAYMENT;
                ViewBag.advanceRefund = v.ADVANCE_PAYMENT_REFUND;
                ViewBag.taxRatio = v.TAX_RATIO;
                ViewBag.retention = v.RETENTION_PAYMENT;
                ViewBag.remark = v.REMARK;
                ViewBag.creatId = v.CREATE_ID;
                ViewBag.createDate = v.CREATE_DATE;
                ViewBag.modifyDate = v.MODIFY_DATE;
                ViewBag.formid = formid;
            }
            return View();
        }

        public String updateVAItem(PLAN_VALUATION_FORM vf, HttpPostedFileBase file)
        {
            logger.Info("create valuation form process! project =" + Request["projectId"]);
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            string msg = "新增計價資料成功!!";
            string Remsg = "修改計價資料成功!!";
            //1.更新或新增業主計價資料
            if (Request["va_amount"] != "")
            {
                vf.VALUATION_AMOUNT = decimal.Parse(Request["va_amount"]);
            }
            if (Request["advance_payment"] != "")
            {
                vf.ADVANCE_PAYMENT = decimal.Parse(Request["advance_payment"]);
            }
            if (Request["tax_ratio"] != "")
            {
                vf.TAX_RATIO = decimal.Parse(Request["tax_ratio"]);
            }
            if (Request["advance_refund"] != "")
            {
                vf.ADVANCE_PAYMENT_REFUND = decimal.Parse(Request["advance_refund"]);
            }
            if (Request["retention_amount"] != "")
            {
                vf.RETENTION_PAYMENT = decimal.Parse(Request["retention_amount"]);
            }
            vf.REMARK = Request["remark"];
            vf.PROJECT_ID = Request["projectId"];
            //新增業主計價資料
            if (null == Request["formid"] || Request["formid"] == "")
            {
                vf.CREATE_ID = uInfo.USER_ID;
                vf.CREATE_DATE = DateTime.Now;
            }
            else
            {
                //修改業主計價資料
                vf.MODIFY_DATE = DateTime.Now;
                vf.CREATE_DATE = Convert.ToDateTime(Request["createDate"]);
                vf.CREATE_ID = Request["creatId"];
            }
            string fid = service.refreshVA(Request["formid"], vf);
            //若使用者有上傳計價附檔，則增加檔案資料
            if (null != file && file.ContentLength != 0)
            {
                //TND_FILE saveF = new TND_FILE();
                //2.解析檔案
                logger.Info("Parser file data:" + file.FileName);
                //2.1 設定檔案名稱,實體位址,附檔名
                string projectid = Request["projectid"];
                string keyName = null;
                if (null != Request["formid"] && Request["formid"] != "")
                {
                    keyName = Request["formid"];
                }
                else
                {
                    keyName = fid;
                }
                logger.Info("file upload namme =" + keyName);
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + Request["projectId"], fileName);
                //saveF.FILE_ACTURE_NAME = fileName;
                var fileType = Path.GetExtension(file.FileName);
                //f.FILE_LOCATIOM = path;
                string createDate = DateTime.Now.ToString("yyyy/MM/dd");
                logger.Info("createDate = " + createDate);
                string createId = uInfo.USER_ID;
                FileManage fs = new FileManage();
                string k = fs.addFile(projectid, keyName, fileName, fileType, path, createId, createDate);
                //int j = service.refreshVAFile(saveF);
                //2.2 將上傳檔案存檔
                logger.Info("save upload file:" + path);
                file.SaveAs(path);
            }
            // 先刪除原先資料
            if (null != Request["formid"] && Request["formid"] != "")
            {
                logger.Info("EST form id =" + Request["formid"]);
                logger.Info("Delete PLAN_INVOICE By EST_FORM_ID");
                service.delInvoiceByESTId(Request["formid"]);
            }
            // 取得憑證資料
            string[] lstDate = Request["invoice_date"].Split(',');
            string[] lstNumber = Request["invoice_number"].Split(',');
            string[] lstAmount = Request["input_amount"].Split(',');
            string[] lstTax = Request["taxamount"].Split(',');
            string[] lstType = Request["invoicetype"].Split(',');
            string[] lstSubType = Request["sub_type"].Split(',');
            string[] lstPlanItem = Request["plan_item_id"].Split(',');
            string[] lstDiscountQty = Request["discount_qty"].Split(',');
            string[] lstDiscountPrice = Request["discount_unit_price"].Split(',');
            List<PLAN_INVOICE> lstItem = new List<PLAN_INVOICE>();
            for (int j = 0; j < lstAmount.Count(); j++)
            {
                PLAN_INVOICE item = new PLAN_INVOICE();
                if (null != Request["formid"] && Request["formid"] != "")
                {
                    item.EST_FORM_ID = Request["formid"];
                }
                else
                {
                    item.EST_FORM_ID = fid;
                }
                item.CONTRACT_ID = Request["projectId"];
                if (lstNumber[j].ToString() == "")
                {
                    item.INVOICE_NUMBER = null;
                }
                else
                {
                    item.INVOICE_NUMBER = lstNumber[j];
                }
                try
                {
                    item.INVOICE_DATE = Convert.ToDateTime(lstDate[j]);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.StackTrace);
                }
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                if (lstTax[j].ToString() == "")
                {
                    item.TAX = null;
                }
                else
                {
                    item.TAX = decimal.Parse(lstTax[j]);
                }
                if (lstType[j].ToString() == "")
                {
                    item.TYPE = null;
                }
                else
                {
                    item.TYPE = lstType[j];
                }
                if (lstType[j] == "折讓單")
                {
                    if (lstSubType[j].ToString() == "")
                    {
                        item.SUB_TYPE = null;
                    }
                    else
                    {
                        item.SUB_TYPE = lstSubType[j];
                    }
                    item.PLAN_ITEM_ID = lstPlanItem[j];
                    if (lstDiscountQty[j].ToString() == "")
                    {
                        item.DISCOUNT_QTY = null;
                    }
                    else
                    {
                        item.DISCOUNT_QTY = decimal.Parse(lstDiscountQty[j]);
                    }
                    if (lstDiscountPrice[j].ToString() == "")
                    {
                        item.DISCOUNT_UNIT_PRICE = null;
                    }
                    else
                    {
                        item.DISCOUNT_UNIT_PRICE = decimal.Parse(lstDiscountPrice[j]);
                    }
                }
                else
                {
                    item.SUB_TYPE = null;
                    item.PLAN_ITEM_ID = null;
                    item.DISCOUNT_QTY = null;
                    item.DISCOUNT_UNIT_PRICE = null;
                }
                logger.Info("Invoice Number = " + item.INVOICE_NUMBER + "and Invoice Amount =" + item.AMOUNT);
                //logger.Debug("Item EST form id =" + item.EST_FORM_ID + "且憑證類型為" + item.TYPE);
                lstItem.Add(item);
            }
            int i = service.addInvoice(lstItem);
            if (fid == "" || null == fid) { msg = service.message; }
            if (null != Request["formid"] && Request["formid"] != "")
            {
                return Remsg;
            }
            else
            {
                return msg;
            }
        }
        /// <summary>
        /// 取得業主計價次數
        /// </summary>
        /// <param name="formid"></param>
        /// <returns></returns>
        public string getAROfForm(string formid)
        {
            logger.Info("get AR of VA form by form id=" + formid);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(service.getVADetailByVAId(formid));
            logger.Info("VA AR's Info=" + itemJson);
            return itemJson;
        }
        //新增應收帳款支付資料
        public String addPaymentDate(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "新增支付資料成功!!";
            PLAN_ACCOUNT item = new PLAN_ACCOUNT();
            FIN_LOAN_TRANACTION loan = new FIN_LOAN_TRANACTION();
            item.PROJECT_ID = form["projectid"];
            item.CONTRACT_ID = form["projectid"];
            item.ACCOUNT_FORM_ID = form["va_form_id"];
            if (form["payment_date"] != "")
            {
                item.PAYMENT_DATE = Convert.ToDateTime(form.Get("payment_date"));
            }
            string loanRemark = "備償款";
            if (form["loan_remark"] != "")
            {
                loanRemark = "備償款-" + form["loan_remark"];
            }
            DateTime paybackDate = DateTime.Now;
            if (form["payment_date"] != "")
            {
                paybackDate = Convert.ToDateTime(form.Get("payment_date"));
            }
            decimal paybackRatio = 0;
            decimal paybackAtm = 0;
            decimal loanBalance = 0;
            decimal bankingFee = 0;
            int period = 0;
            if (null != form.Get("loans").Trim() && form.Get("loans").Trim() != "")
            {
                FIN_BANK_LOAN bl = service.getPaybackRatioById(int.Parse(form.Get("loans").Trim()));
                loanBalance = service.getLoanBalanceByBlId(int.Parse(form.Get("loans").Trim()));
                paybackRatio = decimal.Parse(bl.AR_PAYBACK_RATIO.ToString());
                period = service.getPaybackCountByBlId(int.Parse(form.Get("loans").Trim()));
                if (loanBalance < 0 && loanBalance * -1 < Math.Round(decimal.Parse(form["payment_amount"]) * paybackRatio / 100))
                {
                    paybackAtm = loanBalance * -1;
                }
                else if (loanBalance >= 0)
                {
                    return "此備償戶借款已償還完畢，請重新入帳。";
                }
                else
                {
                    paybackAtm = Math.Round(decimal.Parse(form["payment_amount"]) * paybackRatio / 100);
                }
                logger.Info("paybackAtm = " + paybackAtm);
            }
            if (form["fee"] != "")
            {
                //item.BANKING_FEE = decimal.Parse(form["fee"]);
                bankingFee = decimal.Parse(form["fee"]);
            }
            if (form["payment_amount"] != "")
            {
                item.AMOUNT_PAID = decimal.Parse(form["payment_amount"]) - paybackAtm - bankingFee;
                item.AMOUNT_PAYABLE = decimal.Parse(form["payment_amount"]) - paybackAtm - bankingFee + paybackAtm + bankingFee;
            }
            item.CHECK_NO = form["check_no"];
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            item.CREATE_ID = uInfo.USER_ID;
            item.ISDEBIT = "Y";
            item.ACCOUNT_TYPE = "R";
            item.STATUS = 10;//已支付
            int i = service.addPlanAccount(item);
            if (form.Get("loans").Trim() != "")
            {
                int k = service.addLoanTransaction(int.Parse(form.Get("loans").Trim()), paybackAtm, paybackDate, uInfo.USER_ID, period, loanRemark, form["va_form_id"]);
            }
            if (i == 0) { msg = service.message; }
            return msg;
        }
        //業主計價附檔明細
        public ActionResult VAFileList(string id)
        {
            logger.Info("Access to the Page abount files of Valuation!!");
            List<RevenueFromOwner> lstItem = service.getVAFileByFormId(id);
            logger.Debug("va_form_id = " + id);
            return View(lstItem);
        }
        /// <summary>
        ///  業主計價附檔
        /// </summary>
        public FileResult downLoadVAFile()
        {
            //要下載的檔案位置與檔名
            //string filepath = Request["itemUid"];
            //取得檔案名稱
            FileManage fs = new FileManage();
            TND_FILE f = fs.getFileByItemId(long.Parse(Request["itemid"]));
            string filename = System.IO.Path.GetFileName(f.FILE_LOCATIOM);
            //讀成串流
            Stream iStream = new FileStream(f.FILE_LOCATIOM, FileMode.Open, FileAccess.Read, FileShare.Read);
            //回傳出檔案
            return File(iStream, "application/unknown", filename);
            // return File(iStream, "application/zip", filename);//application/unknown
        }

        //刪除單一附檔資料
        public String delVAFile()
        {
            long itemUid = long.Parse(Request["itemid"]);
            SYS_USER loginUser = (SYS_USER)Session["user"];
            logger.Info(loginUser.USER_ID + " remove data:va file uid=" + itemUid);
            FileManage fs = new FileManage();
            TND_FILE f = fs.getFileByItemId(long.Parse(Request["itemid"]));
            var path = Path.Combine(ContextService.strUploadPath + "/" + f.PROJECT_ID, f.FILE_ACTURE_NAME);
            System.IO.File.Delete(path);
            int i = fs.delFile(itemUid);
            return "檔案已刪除(" + i + ")";
        }

        //上傳廠商計價檔案
        public String uploadFile4Supplier(HttpPostedFileBase file)
        {
            string msg = "新增檔案成功!!";
            string k = null;
            //若使用者有上傳檔案，則增加檔案資料
            if (null != file && file.ContentLength != 0)
            {
                //2.解析檔案
                logger.Info("Parser file data:" + file.FileName);
                //2.1 設定檔案名稱,實體位址,附檔名
                string projectid = Request["projectId"];
                string keyName = Request["uploadName"];
                logger.Info("file upload namme =" + keyName);
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "\\" + Request["projectId"], fileName);
                var fileType = Path.GetExtension(file.FileName);
                string createDate = DateTime.Now.ToString("yyyy/MM/dd");
                logger.Info("createDate = " + createDate);
                UserService us = new UserService();
                SYS_USER u = (SYS_USER)Session["user"];
                SYS_USER uInfo = us.getUserInfo(u.USER_ID);
                string createId = uInfo.USER_ID;
                FileManage fs = new FileManage();
                k = fs.addFile(projectid, keyName, fileName, fileType, path, createId, createDate);
                //2.2 將上傳檔案存檔
                logger.Info("save upload file:" + path);
                file.SaveAs(path);
            }
            if (k == "" || null == k) { msg = service.message; }
            return msg;
        }
        /// <summary>
        /// 下載折讓單
        /// </summary>
        public void downLoadCreditNote()
        {
            string formid = Request["formid"];
            string projectid = Request["projectid"];
            List<CreditNote> lstInvoice = service.getCreditNoteById(projectid, formid);
            RevenueFromOwner va = service.getVADetailByVAId(formid);
            if (lstInvoice.Count > 0)
            {
                CreditNoteToExcel poi = new CreditNoteToExcel();
                //檔案位置
                string fileLocation = poi.exportExcel(lstInvoice, va);
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
using log4net;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;
using System.IO;
using System.Web.Script.Serialization;
using Newtonsoft.Json;


namespace topmeperp.Controllers
{
    public class CashFlowController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        Service4CashFlow service = new Service4CashFlow();

        // GET: CashFlow 
        [topmeperp.Filter.AuthFilter]
        public ActionResult Index()
        {
            List<ProjectList> lstProject = PlanService.SearchProjectByName("", "專案執行','保固",null);
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            return View(lstProject);
        }

        public ActionResult CashFlowManage()
        {
            List<CashFlowFunction> lstCashFlow = null;
            List<PlanFinanceProfile> lstFinProfile = null;
            CashFlowBalance cashFlowBalance = null;
            lstCashFlow = service.getCashFlow();
            lstFinProfile = service.getPlanFinProfile();
            //totalFinProfile = service.getFinProfile();
            cashFlowBalance = service.getCashFlowBalance();
            CashFlowModel viewModel = new CashFlowModel();
            viewModel.finFlow = lstCashFlow;
            viewModel.finProfile = lstFinProfile;
            //todo
            //viewModel.totalFinProfile = totalFinProfile;
            viewModel.finBalance = cashFlowBalance;
            ViewBag.today = DateTime.Now;
            return View(viewModel);
        }
        //取得特定日期收入明細
        public ActionResult CashInFlowItem(string paymentDate, string type)
        {
            List<PlanAccountFunction> CashInFlow = null;
            List<LoanTranactionFunction> LoanInFlow = null;
            CashFlowModel viewModel = new CashFlowModel();
            string projectname = "";
            string account_type = "R";
            int day = (Convert.ToDateTime(paymentDate) - Convert.ToDateTime("2018/01/01")).Days % 7;
            string lstDate = Convert.ToDateTime(paymentDate).AddDays(-2).ToString("yyyy/MM/dd");
            int dayOverHalfAYear = (Convert.ToDateTime(paymentDate) - DateTime.Now).Days;
            string DateOverHalfAYear = DateTime.Now.AddDays(181).ToString("yyyy/MM/dd");
            logger.Debug("day's mod =" + day);
            if (dayOverHalfAYear > 180)
            {
                CashInFlow = service.getPlanAccount(null, projectname, projectname, account_type, projectname, DateOverHalfAYear, paymentDate);
                LoanInFlow = service.getLoanTranaction(type, null, DateOverHalfAYear, paymentDate);

            }
            else if (day == 0)
            {
                CashInFlow = service.getPlanAccount(null, projectname, projectname, account_type, projectname, lstDate, paymentDate);
                LoanInFlow = service.getLoanTranaction(type, null, lstDate, paymentDate);
            }
            else
            {
                CashInFlow = service.getPlanAccount(paymentDate, projectname, projectname, account_type, projectname, null, null);
                LoanInFlow = service.getLoanTranaction(type, paymentDate, null, null);
            }
            viewModel.planAccount = CashInFlow;
            viewModel.finLoanTranaction = LoanInFlow;
            ViewBag.SearchTerm = "實際入帳日期為 : " + paymentDate;
            return View(viewModel);
        }

        //取得特定日期支出明細
        public ActionResult CashOutFlowItem(string paymentDate, string type)
        {
            List<PlanAccountFunction> CashOutFlow = null; //廠商請款
            List<LoanTranactionFunction> LoanOutFlow = null;//借款還款 / 廠商借款
            List<PlanAccountFunction> OutFlowBalance = null;//費用支出
            List<ExpenseFormFunction> ExpenseOutFlow = null;
            List<ExpenseFormFunction> ExpenseBudget = null;

            CashFlowModel viewModel = new CashFlowModel();
            string projectname = "";
            string account_type = "P";
            int day = (Convert.ToDateTime(paymentDate) - Convert.ToDateTime("2018/01/01")).Days % 7;
            string lstDate = Convert.ToDateTime(paymentDate).AddDays(-2).ToString("yyyy/MM/dd");
            int dayOverHalfAYear = (Convert.ToDateTime(paymentDate) - DateTime.Now).Days;
            string DateOverHalfAYear = DateTime.Now.AddDays(181).ToString("yyyy/MM/dd");
            logger.Debug("day's mod =" + day);
            if (dayOverHalfAYear > 180)
            {
                CashOutFlow = service.getPlanAccount(null, projectname, projectname, account_type, projectname, DateOverHalfAYear, paymentDate);
                LoanOutFlow = service.getLoanTranaction(type, null, DateOverHalfAYear, paymentDate);
                OutFlowBalance = service.getOutFlowBalanceByDate(null, DateOverHalfAYear, paymentDate);
                ExpenseOutFlow = service.getExpenseOutFlowByDate(null, DateOverHalfAYear, paymentDate);
                ExpenseBudget = service.getExpenseBudgetByDate(null, DateOverHalfAYear, paymentDate);
            }

            else if (day == 0)
            {
                CashOutFlow = service.getPlanAccount(null, projectname, projectname, account_type, projectname, lstDate, paymentDate);
                LoanOutFlow = service.getLoanTranaction(type, null, lstDate, paymentDate);
                OutFlowBalance = service.getOutFlowBalanceByDate(null, lstDate, paymentDate);
                ExpenseOutFlow = service.getExpenseOutFlowByDate(null, lstDate, paymentDate);
                ExpenseBudget = service.getExpenseBudgetByDate(null, lstDate, paymentDate);
            }
            else
            {
                CashOutFlow = service.getPlanAccount(paymentDate, projectname, projectname, account_type, projectname, null, null);
                LoanOutFlow = service.getLoanTranaction(type, paymentDate, null, null);
                OutFlowBalance = service.getOutFlowBalanceByDate(paymentDate, null, null);
                ExpenseOutFlow = service.getExpenseOutFlowByDate(paymentDate, null, null);
                ExpenseBudget = service.getExpenseBudgetByDate(paymentDate, null, null);
            }
            //廠商請款(計價單) 
            viewModel.planAccount = CashOutFlow;
            //借款還款/廠商借款 
            viewModel.finLoanTranaction = LoanOutFlow;
            //費用單
            viewModel.outFlowBalance = OutFlowBalance;
            //費用預算(含公司與工地，不含今天以前的資料)
            viewModel.expBudget = ExpenseBudget;
            //當日須支付明細
            viewModel.outFlowExp = ExpenseOutFlow;
            ViewBag.SearchTerm = "實際付款日期為 : " + paymentDate;
            return View(viewModel);
        }

        //查詢公司預算
        public ActionResult Search()
        {
            List<ExpenseBudgetSummary> ExpBudget = null;
            //List<ExpenseBudgetByMonth> BudgetByMonth = null;
            ExpenseBudgetModel viewModel = new ExpenseBudgetModel();
            ExpenseBudgetSummary Amt = null;
            if (null != Request["budgetyear"])
            {
                //年度預算月分配
                ExpBudget = service.getExpBudgetByYear(int.Parse(Request["budgetyear"]));
                //年度預算月分配總和
                //BudgetByMonth = service.getExpBudgetOfMonthByYear(int.Parse(Request["budgetyear"]));
                Amt = service.getTotalExpBudgetAmount(int.Parse(Request["budgetyear"]));
                viewModel.BudgetSummary = ExpBudget;
                //viewModel.budget = BudgetByMonth;
                TempData["TotalAmt"] = Amt.TOTAL_BUDGET;
                TempData["budgetYear"] = Request["budgetyear"];
                return View("ExpenseBudget", viewModel);
            }
            TempData["budgetYear"] = Request["budgetyear"];
            return View("ExpenseBudget");
        }
        /// <summary>
        /// 下載公司費用預算填寫表
        /// </summary>
        public void downLoadExpBudgetForm()
        {
            ExpBudgetFormToExcel poi = new ExpBudgetFormToExcel();
            //檔案位置
            string fileLocation = poi.exportExcel();
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
        //上傳公司費用預算
        [HttpPost]
        public ActionResult uploadExpBudgetTable(HttpPostedFileBase fileBudget)
        {
            int budgetYear = int.Parse(Request["year"]);
            logger.Info("Upload Expense Budget Table for budget year =" + budgetYear);
            string message = "";

            if (null != fileBudget && fileBudget.ContentLength != 0)
            {
                //2.解析Excel
                logger.Info("Parser Excel data:" + fileBudget.FileName);
                //2.1 設定Excel 檔案名稱
                var fileName = Path.GetFileName(fileBudget.FileName);
                var path = Path.Combine(ContextService.strUploadPath, fileName);
                logger.Info("save excel file:" + path);
                fileBudget.SaveAs(path);
                //2.2 開啟Excel 檔案
                logger.Info("Parser Excel File Begin:" + fileBudget.FileName);
                ExpBudgetFormToExcel budgetservice = new ExpBudgetFormToExcel();
                budgetservice.InitializeWorkbook(path);
                //解析預算數量
                List<FIN_EXPENSE_BUDGET> lstExpBudget = budgetservice.ConvertDataForExpBudget(budgetYear);
                //2.3 記錄錯誤訊息
                message = budgetservice.errorMessage;
                //2.4
                logger.Info("Delete FIN_EXPENSE_BUDGET By Year");
                service.delExpBudgetByYear(budgetYear);
                message = message + "<br/>舊有資料刪除成功 !!";
                //2.5 
                logger.Info("Add All FIN_EXPENSE_BUDGET to DB");
                service.refreshExpBudget(lstExpBudget);
                message = message + "<br/>資料匯入完成 !!";
            }
            TempData["result"] = message;
            return RedirectToAction("ExpenseBudget");
        }

        public ActionResult ExpenseBudget()
        {
            logger.Info("Access to Expense Budget Page !!");
            return View();
        }
        //更新公司費用預算
        public String UpdateExpBudget(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            // 先刪除原先資料
            logger.Info("Expense Budget Year =" + form["year"]);
            logger.Info("Delete FIN_EXPENSE_BUDGET By BUDGET_YEAR");
            service.delExpBudgetByYear(int.Parse(form["year"]));
            string msg = "";
            string[] lstsubjectid = form.Get("subjectid").Split(',');
            string[] lst7 = new string[lstsubjectid.Length];
            for (int i = 0; i < lstsubjectid.Length; i++)
            {
                lst7[i] = form.Get("julAmt" + lstsubjectid[i]).Replace(",", "");
                logger.Debug("get budget jul=" + lst7[i]);
            }
            string[] lst8 = new string[lstsubjectid.Length];
            for (int i = 0; i < lstsubjectid.Length; i++)
            {
                lst8[i] = form.Get("augAmt" + lstsubjectid[i]).Replace(",", "");
                logger.Debug("get budget aug=" + lst8[i]);
            }
            string[] lst9 = new string[lstsubjectid.Length];
            for (int i = 0; i < lstsubjectid.Length; i++)
            {
                lst9[i] = form.Get("sepAmt" + lstsubjectid[i]).Replace(",", "");
                logger.Debug("get budget sep=" + lst9[i]);
            }
            string[] lst10 = new string[lstsubjectid.Length];
            for (int i = 0; i < lstsubjectid.Length; i++)
            {
                lst10[i] = form.Get("octAmt" + lstsubjectid[i]).Replace(",", "");
                logger.Debug("get budget oct=" + lst10[i]);
            }
            string[] lst11 = new string[lstsubjectid.Length];
            for (int i = 0; i < lstsubjectid.Length; i++)
            {
                lst11[i] = form.Get("novAmt" + lstsubjectid[i]).Replace(",", "");
                logger.Debug("get budget nov=" + lst11[i]);
            }
            string[] lst12 = new string[lstsubjectid.Length];
            for (int i = 0; i < lstsubjectid.Length; i++)
            {
                lst12[i] = form.Get("decAmt" + lstsubjectid[i]).Replace(",", "");
                logger.Debug("get budget dec=" + lst12[i]);
            }
            string[] lst1 = new string[lstsubjectid.Length];
            for (int i = 0; i < lstsubjectid.Length; i++)
            {
                lst1[i] = form.Get("janAmt" + lstsubjectid[i]).Replace(",", "");
                logger.Debug("get budget jan=" + lst1[i]);
            }
            string[] lst2 = new string[lstsubjectid.Length];
            for (int i = 0; i < lstsubjectid.Length; i++)
            {
                lst2[i] = form.Get("febAmt" + lstsubjectid[i]).Replace(",", "");
                logger.Debug("get budget feb=" + lst2[i]);
            }
            string[] lst3 = new string[lstsubjectid.Length];
            for (int i = 0; i < lstsubjectid.Length; i++)
            {
                lst3[i] = form.Get("marAmt" + lstsubjectid[i]).Replace(",", "");
                logger.Debug("get budget mar=" + lst3[i]);
            }
            string[] lst4 = new string[lstsubjectid.Length];
            for (int i = 0; i < lstsubjectid.Length; i++)
            {
                lst4[i] = form.Get("aprAmt" + lstsubjectid[i]).Replace(",", "");
                logger.Debug("get budget apr=" + lst4[i]);
            }
            string[] lst5 = new string[lstsubjectid.Length];
            for (int i = 0; i < lstsubjectid.Length; i++)
            {
                lst5[i] = form.Get("mayAmt" + lstsubjectid[i]).Replace(",", "");
                logger.Debug("get budget may=" + lst5[i]);
            }
            string[] lst6 = new string[lstsubjectid.Length];
            for (int i = 0; i < lstsubjectid.Length; i++)
            {
                lst6[i] = form.Get("junAmt" + lstsubjectid[i]).Replace(",", "");
                logger.Debug("get budget jun=" + lst6[i]);
            }

            List<string[]> Atm = new List<string[]>();
            Atm.Add(lst7);
            Atm.Add(lst8);
            Atm.Add(lst9);
            Atm.Add(lst10);
            Atm.Add(lst11);
            Atm.Add(lst12);
            Atm.Add(lst1);
            Atm.Add(lst2);
            Atm.Add(lst3);
            Atm.Add(lst4);
            Atm.Add(lst5);
            Atm.Add(lst6);
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            List<FIN_EXPENSE_BUDGET> lst = new List<FIN_EXPENSE_BUDGET>();
            for (int j = 0; j < lstsubjectid.Count(); j++)
            {
                List<FIN_EXPENSE_BUDGET> lstItem = new List<FIN_EXPENSE_BUDGET>();
                for (int i = 0; i < 6; i++)
                {
                    FIN_EXPENSE_BUDGET item = new FIN_EXPENSE_BUDGET();
                    item.BUDGET_YEAR = int.Parse(form["year"]);
                    item.SUBJECT_ID = lstsubjectid[j];
                    item.MODIFY_ID = u.USER_ID;
                    item.CURRENT_YEAR = int.Parse(form["year"]);
                    if (Atm[i][j].ToString() == "" || null == Atm[i][j].ToString())
                    {
                        item.AMOUNT = null;
                        item.BUDGET_MONTH = i + 7;
                    }
                    else
                    {
                        item.AMOUNT = decimal.Parse(Atm[i][j]);
                        item.BUDGET_MONTH = i + 7;
                    }
                    item.MODIFY_DATE = DateTime.Now;
                    logger.Info("費用項目代碼 =" + item.SUBJECT_ID + "，and Budget Month = " + item.BUDGET_MONTH + "，and Atm = " + Atm[i][j]);
                    lstItem.Add(item);
                }
                for (int i = 6; i < 12; i++)
                {
                    FIN_EXPENSE_BUDGET item = new FIN_EXPENSE_BUDGET();
                    item.BUDGET_YEAR = int.Parse(form["year"]);
                    item.SUBJECT_ID = lstsubjectid[j];
                    item.MODIFY_ID = u.USER_ID;
                    item.CURRENT_YEAR = int.Parse(form["year"]) + 1;
                    if (Atm[i][j].ToString() == "" || null == Atm[i][j].ToString())
                    {
                        item.AMOUNT = null;
                        item.BUDGET_MONTH = i - 5;
                    }
                    else
                    {
                        item.AMOUNT = decimal.Parse(Atm[i][j]);
                        item.BUDGET_MONTH = i - 5;
                    }
                    item.MODIFY_DATE = DateTime.Now;
                    logger.Info("費用項目代碼 =" + item.SUBJECT_ID + "，and Budget Month = " + item.BUDGET_MONTH + "，and Atm = " + Atm[i][j]);
                    lstItem.Add(item);
                }
                lst.AddRange(lstItem);
            }
            int k = service.refreshExpBudget(lst);
            if (k == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新公司預算費用成功，預算年度為 " + form["year"];
            }

            logger.Info("Request:BUDGET_YEAR =" + form["year"]);
            return msg;
        }
        //申請公司費用
        public ActionResult OperatingExpense()
        {
            logger.Info("Access to Operating Expense Page !!");
            List<FIN_SUBJECT> Subject = null;
            Subject = service.getSubjectOfExpense();
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
            List<FIN_SUBJECT> SubjectChecked = null;
            SubjectChecked = service.getSubjectByChkItem(lstItemId);
            List<FIN_SUBJECT> Subject = null;
            Subject = service.getSubjectOfExpense();
            ViewData["items"] = JsonConvert.SerializeObject(Subject);
            return View("OperatingExpense", SubjectChecked);
        }
        //單張費用單中新增的項目個數必須<= 58,以便套用到form_expense Template 的格式設定
        [HttpPost]
        public ActionResult AddExpense(FIN_EXPENSE_FORM ef, FormCollection form)
        {
            //新增公司費用申請單
            logger.Info("form:" + form.Count);
            string[] lstSubject = form.Get("subjectid").Split(',');
            //可處理千分位符號!!
            string[] lstAmount = (string[])form.GetValue("input_amount").RawValue;
            string[] lstPrice = (string[])form.GetValue("unit_price").RawValue;
            string[] lstRemark = form.Get("item_remark").Split(',');
            string[] lstUnit = form.Get("unit").Split(',');
            string[] lstQty = form.Get("item_quantity").Split(',');
            string[] SubjectList = form.Get("subjectlist").Split(',');
            logger.Debug("SubjectList = " + SubjectList);
            //建立公司/工地費用單號
            logger.Info("create new Operating Expense Form");
            if (null != form["projectid"] || form["projectid"] != "")
            {
                ef.PROJECT_ID = form["projectid"];
            }
            SYS_USER uInfo = (SYS_USER)Session["user"];
            if (null != Request["paymentdate"] && "" != Request["paymentdate"])
            {
                ef.PAYMENT_DATE = Convert.ToDateTime(Request["paymentdate"]);
            }
            ef.OCCURRED_YEAR = int.Parse(Request["paymentdate"].Substring(0, 4));
            ef.OCCURRED_MONTH = int.Parse(Request["paymentdate"].Substring(5, 2));
            ef.CREATE_DATE = DateTime.Now;
            ef.CREATE_ID = uInfo.USER_ID;
            ef.REMARK = Request["remark"];
            ef.PAYEE = Request["supplier"];
            ef.STATUS = 10;
            string fid = service.newExpenseForm(ef);
            //建立公司/工費用單明細
            List<FIN_EXPENSE_ITEM> lstItem = new List<FIN_EXPENSE_ITEM>();
            for (int j = 0; j < lstSubject.Count(); j++)
            {
                FIN_EXPENSE_ITEM item = new FIN_EXPENSE_ITEM();
                item.FIN_SUBJECT_ID = SubjectList[int.Parse(lstSubject[j])];
                item.ITEM_REMARK = lstRemark[j];
                item.ITEM_UNIT = lstUnit[j];
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                if (lstPrice[j].ToString() == "")
                {
                    item.ITEM_UNIT_PRICE = null;
                }
                else
                {
                    item.ITEM_UNIT_PRICE = decimal.Parse(lstPrice[j]);
                }
                if (lstQty[j].ToString() == "")
                {
                    item.ITEM_QUANTITY = null;
                }
                else
                {
                    item.ITEM_QUANTITY = decimal.Parse(lstQty[j]);
                }
                logger.Info("Operating Expense Subject =" + item.FIN_SUBJECT_ID + "， and Amount = " + item.AMOUNT);
                item.EXP_FORM_ID = fid;
                logger.Debug("Item EX form id =" + item.EXP_FORM_ID);
                lstItem.Add(item);
            }
            int i = service.AddExpenseItems(lstItem);
            //建立公司費用申請參考流程
            if (null == form["projectid"] || form["projectid"] == "")
            {
                Flow4CompanyExpense flowService = new Flow4CompanyExpense();
                logger.Debug("Item Count =" + i);
                flowService.iniRequest(uInfo, fid);
            }
            else
            {
                //建立工地費用申請參考流程
                Flow4SiteExpense flowService = new Flow4SiteExpense();
                logger.Debug("Item Count =" + i);
                flowService.iniRequest(uInfo, fid);

            }
            return Redirect("SingleEXPForm?id=" + fid);
        }

        //顯示單一公司營業費用單/工地費用單功能
        public ActionResult SingleEXPForm(string id)
        {
            logger.Info("http get mehtod:" + id);
            OperatingExpenseModel singleForm = new OperatingExpenseModel();
            Flow4CompanyExpense wfs = new Flow4CompanyExpense();

            service.getEXPByExpId(id);
            singleForm.finEXP = service.formEXP;
            singleForm.finEXPItem = service.EXPItem;
            singleForm.planEXPItem = service.siteEXPItem;

            logger.Info("get process request by dataId=" + id);
            wfs.getTask(id);
            wfs.getRequest(id);
            wfs.task.FormData = singleForm;

            Session["process"] = wfs.task;
            return View(wfs.task);
        }
        //送審、通過
        public String SendForm(FormCollection f)
        {
            logger.Info("http get mehtod:" + f["EXP_FORM_ID"]);
            Flow4CompanyExpense wfs = new Flow4CompanyExpense();
            wfs.task = (ExpenseTask)Session["process"];
            logger.Info("Data In Session :" + wfs.task.FormData.finEXP.EXP_FORM_ID);

            SYS_USER u = (SYS_USER)Session["user"];
            DateTime? paymentdate = null;//DateTime can not set null
            string desc = null;
            if (f["paymentdate"].ToString() != "")
            {
                paymentdate = Convert.ToDateTime(f["paymentdate"].ToString());
            }
            if (null != f["RejectDesc"] && f["RejectDesc"].ToString() != "")
            {
                desc = f["RejectDesc"].ToString().Trim();
            }

            wfs.Send(u, paymentdate, desc);
            return "更新成功!!";
        }
        //退件
        public String RejectForm(FormCollection form)
        {
            //取得表單資料 from Session
            Flow4CompanyExpense wfs = new Flow4CompanyExpense();
            wfs.task = (ExpenseTask)Session["process"];
            SYS_USER u = (SYS_USER)Session["user"];
            wfs.Reject(u, null, form["RejectDesc"]);
            return wfs.Message;
        }
        //取消
        public String CancelForm(FormCollection form)
        {
            Flow4CompanyExpense wfs = new Flow4CompanyExpense();
            wfs.task = (ExpenseTask)Session["process"];
            SYS_USER u = (SYS_USER)Session["user"];
            wfs.Cancel(u);
            return wfs.Message;
        }
        //更新費用單
        public String UpdateEXP(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            // 取得費用單資料
            FIN_EXPENSE_FORM ef = new FIN_EXPENSE_FORM();
            ef.OCCURRED_YEAR = int.Parse(form.Get("year").Trim());
            ef.OCCURRED_MONTH = int.Parse(form.Get("month").Trim());
            ef.REMARK = form.Get("remark").Trim();
            ef.CREATE_ID = form.Get("createid").Trim();
            if ("" != form.Get("paymentdate"))
            {
                ef.PAYMENT_DATE = Convert.ToDateTime(form.Get("paymentdate"));
            }
            ef.CREATE_DATE = Convert.ToDateTime(form.Get("createdate"));
            ef.STATUS = int.Parse(form.Get("status").Trim());
            ef.MODIFY_DATE = DateTime.Now;
            ef.EXP_FORM_ID = form.Get("formnumber").Trim();
            ef.PROJECT_ID = form.Get("projectid").Trim();
            ef.PAYEE = form.Get("supplier").Trim();
            string[] lstSubject = form.Get("subject").Split(',');
            string[] lstRemark = form.Get("item_remark").Split(',');
            string[] lstUnit = form.Get("unit").Split(',');
            string[] lstQty = form.Get("item_quantity").Split(',');
            //可處理千分位符號!!
            string[] lstAmount = (string[])form.GetValue("amount").RawValue;
            string[] lstPrice = (string[])form.GetValue("unit_price").RawValue;
            string[] lstExpItemId = form.Get("exp_item_id").Split(',');
            string formid = form.Get("formnumber").Trim();
            List<FIN_EXPENSE_ITEM> lstItem = new List<FIN_EXPENSE_ITEM>();
            for (int j = 0; j < lstSubject.Count(); j++)
            {
                FIN_EXPENSE_ITEM item = new FIN_EXPENSE_ITEM();
                item.FIN_SUBJECT_ID = lstSubject[j];
                item.EXP_ITEM_ID = int.Parse(lstExpItemId[j]);
                if (lstRemark[j].ToString() == "")
                {
                    item.ITEM_REMARK = null;
                }
                else
                {
                    item.ITEM_REMARK = lstRemark[j];
                }
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                if (lstPrice[j].ToString() == "")
                {
                    item.ITEM_UNIT_PRICE = null;
                }
                else
                {
                    item.ITEM_UNIT_PRICE = decimal.Parse(lstPrice[j]);
                }
                if (lstQty[j].ToString() == "")
                {
                    item.ITEM_QUANTITY = null;
                }
                else
                {
                    item.ITEM_QUANTITY = decimal.Parse(lstQty[j]);
                }
                if (lstUnit[j].ToString() == "")
                {
                    item.ITEM_UNIT = null;
                }
                else
                {
                    item.ITEM_UNIT = lstUnit[j];
                }
                logger.Debug("Expense Item Id =" + item.EXP_ITEM_ID + ", Subject Id =" + item.FIN_SUBJECT_ID + ", Amount =" + item.AMOUNT);
                lstItem.Add(item);
            }
            int i = service.refreshEXPForm(formid, ef, lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else if (form["projectid"] != null && form["projectid"] != "")
            {
                msg = "更新工地費用單成功";
            }
            else
            {
                msg = "更新公司營業費用單成功";
            }
            logger.Info("Request: 更新公司營業費用/工地費用單訊息 = " + msg);
            return msg;
        }

        public String UpdateEXPStatusById(FormCollection form)
        {
            //取得費用單編號
            logger.Info("form:" + form.Count);
            logger.Info("EXP form Id:" + form["formnumber"]);
            string msg = "";
            FIN_EXPENSE_FORM ef = new FIN_EXPENSE_FORM();
            ef.OCCURRED_YEAR = int.Parse(form.Get("year").Trim());
            ef.OCCURRED_MONTH = int.Parse(form.Get("month").Trim());
            ef.REMARK = form.Get("remark").Trim();
            ef.CREATE_ID = form.Get("createid").Trim();
            ef.PAYMENT_DATE = Convert.ToDateTime(form.Get("paymentdate"));
            ef.CREATE_DATE = Convert.ToDateTime(form.Get("createdate"));
            ef.STATUS = int.Parse(form.Get("status").Trim());
            ef.MODIFY_DATE = DateTime.Now;
            ef.EXP_FORM_ID = form.Get("formnumber").Trim();
            ef.PROJECT_ID = form.Get("projectid").Trim();
            ef.PAYEE = form.Get("supplier").Substring(0, 7);
            string[] lstSubject = form.Get("subject").Split(',');
            string[] lstRemark = form.Get("item_remark").Split(',');
            string[] lstUnit = form.Get("unit").Split(',');
            string[] lstQty = form.Get("item_quantity").Split(',');
            //可處理千分位符號!!
            string[] lstAmount = (string[])form.GetValue("amount").RawValue;
            string[] lstPrice = (string[])form.GetValue("unit_price").RawValue;
            string[] lstExpItemId = form.Get("exp_item_id").Split(',');
            string formid = form.Get("formnumber").Trim();
            List<FIN_EXPENSE_ITEM> lstItem = new List<FIN_EXPENSE_ITEM>();
            for (int j = 0; j < lstSubject.Count(); j++)
            {
                FIN_EXPENSE_ITEM item = new FIN_EXPENSE_ITEM();
                item.FIN_SUBJECT_ID = lstSubject[j];
                item.EXP_ITEM_ID = int.Parse(lstExpItemId[j]);
                if (lstRemark[j].ToString() == "")
                {
                    item.ITEM_REMARK = null;
                }
                else
                {
                    item.ITEM_REMARK = lstRemark[j];
                }
                if (lstAmount[j].ToString() == "")
                {
                    item.AMOUNT = null;
                }
                else
                {
                    item.AMOUNT = decimal.Parse(lstAmount[j]);
                }
                if (lstPrice[j].ToString() == "")
                {
                    item.ITEM_UNIT_PRICE = null;
                }
                else
                {
                    item.ITEM_UNIT_PRICE = decimal.Parse(lstPrice[j]);
                }
                if (lstQty[j].ToString() == "")
                {
                    item.ITEM_QUANTITY = null;
                }
                else
                {
                    item.ITEM_QUANTITY = decimal.Parse(lstQty[j]);
                }
                if (lstUnit[j].ToString() == "")
                {
                    item.ITEM_UNIT = null;
                }
                else
                {
                    item.ITEM_UNIT = lstUnit[j];
                }
                logger.Debug("Expense Item Id =" + item.EXP_ITEM_ID + ", Subject Id =" + item.FIN_SUBJECT_ID + ", Amount =" + item.AMOUNT);
                lstItem.Add(item);
            }
            int i = service.refreshEXPForm(formid, ef, lstItem);
            //更新費用單狀態
            logger.Info("Update Expense Form Status");
            //費用單(已送審) STATUS = 20
            int k = service.RefreshEXPStatusById(formid);
            if (k == 0)
            {
                msg = service.message;
            }
            else if (form["projectid"] != null && form["projectid"] != "")
            {
                msg = "工地費用單已送審";
            }
            else
            {
                msg = "公司營業費用單已送審";
            }
            return msg;
        }

        //費用單查詢
        public ActionResult ExpenseForm(string id)
        {
            if (id != null && id != "")
            {
                TND_PROJECT p = service.getProjectById(id);
                ViewBag.projectName = p.PROJECT_NAME;
                ViewBag.projectid = id;
            }
            else
            {
                id = "";
                ViewBag.projectid = "";
            }
            //取得表單狀態參考資料
            SelectList status = new SelectList(SystemParameter.getSystemPara("ExpenseForm"), "KEY_FIELD", "VALUE_FIELD");
            ViewData.Add("status", status);
            Flow4CompanyExpense s = new Flow4CompanyExpense();
            List<ExpenseFlowTask> lstEXP = s.getCompanyExpenseRequest(Request["occurred_date"], Request["subjectname"], Request["expid"], id, null);
            return View(lstEXP);
        }

        public ActionResult SearchEXP()
        {
            string id = Request["id"];
            string status = Request["status"];
            if (id != null && id != "")
            {
                TND_PROJECT p = service.getProjectById(id);
                ViewBag.projectName = p.PROJECT_NAME;
                ViewBag.projectid = id;
            }
            SelectList LstStatus = new SelectList(SystemParameter.getSystemPara("ExpenseForm"), "KEY_FIELD", "VALUE_FIELD");
            ViewData.Add("status", LstStatus);
            Flow4CompanyExpense s = new Flow4CompanyExpense();

            List<ExpenseFlowTask> lstEXP = s.getCompanyExpenseRequest(Request["occurred_date"], Request["subjectname"], Request["expid"], id, status);
            ViewBag.SearchResult = "共取得" + lstEXP.Count + "筆資料";
            return View("ExpenseForm", lstEXP);
        }

        public String UpdateAccountStatus(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            int i = 0;
            string[] lstForm = form.Get("plan_account_id").Split(',');
            string[] lstDate = form.Get("payment_date").Split(',');
            string[] lstCheck = form.Get("check_no").Split(',');
            List<PLAN_ACCOUNT> lstItem = new List<PLAN_ACCOUNT>();
            if (form.Get("status") != null)
            {
                string[] lstStatus = form.Get("status").Split(',');
                for (int j = 0; j < lstForm.Count(); j++)
                {
                    PLAN_ACCOUNT item = new PLAN_ACCOUNT();
                    item.PLAN_ACCOUNT_ID = int.Parse(lstForm[j]);
                    item.PAYMENT_DATE = Convert.ToDateTime(lstDate[j]);
                    item.CHECK_NO = lstCheck[j];
                    item.MODIFY_DATE = DateTime.Now;
                    if (lstStatus[j].ToString() == "")
                    {
                        item.STATUS = 10;
                    }
                    else
                    {
                        item.STATUS = 0;
                    }
                    logger.Debug("Plan Acount Id =" + item.PLAN_ACCOUNT_ID + ", Status =" + item.STATUS);
                    lstItem.Add(item);
                }
            }
            else
            {
                for (int j = 0; j < lstForm.Count(); j++)
                {
                    PLAN_ACCOUNT item = new PLAN_ACCOUNT();
                    item.PLAN_ACCOUNT_ID = int.Parse(lstForm[j]);
                    item.PAYMENT_DATE = Convert.ToDateTime(lstDate[j]);
                    item.CHECK_NO = lstCheck[j];
                    item.STATUS = 10;
                    item.MODIFY_DATE = DateTime.Now;
                    logger.Debug("Plan Acount Id =" + item.PLAN_ACCOUNT_ID + ", Status =" + item.STATUS);
                    lstItem.Add(item);
                }
            }
            i = service.refreshAccountStatus(lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "帳款支付狀態已更新";
            }
            return msg;
        }

        public ActionResult FormForJournal()
        {
            logger.Info("Access to Form For Journal !!");
            //公司需立帳之帳款(即會計審核)
            int status = 30;
            List<OperatingExpenseFunction> lstEXP = service.getEXPListByExpId(null, null, null, status, null);
            return View(lstEXP);
        }

        //會計稽核作業
        public ActionResult SearchForm4Journal()
        {
            //logger.Info("occurred_date =" + Request["occurred_date"] + ", subjectname =" + Request["subjectname"] + ", expid =" + Request["expid"] + ", status =" + int.Parse(Request["status"]) + ", projectid =" + Request["id"]);
            List<OperatingExpenseFunction> lstEXP = new List<OperatingExpenseFunction>();//service.getEXPListByExpId(Request["occurred_date"], Request["subjectname"], Request["expid"], int.Parse(Request["status"]), Request["id"]);
            //ViewBag.SearchResult = "共取得" + lstEXP.Count + "筆資料";
            return View("FormForJournal", lstEXP);
        }

        //修改帳款支付日期
        public ActionResult PlanAccount()
        {
            logger.Info("Search For Account To Update Its Payment Date !!");
            return View();
        }

        public ActionResult ShowPlanAccount()
        {
            logger.Info("payment_date =" + Request["payment_date"] + ", projectname =" + Request["projectname"] + ", payee =" + Request["payee"] + ", account_type =" + Request["account_type"]);
            string formId = "";
            List<PlanAccountFunction> lstAccount = service.getPlanAccount(Request["payment_date"], Request["projectname"], Request["payee"], Request["account_type"], formId, null, null);
            ViewBag.SearchResult = "共取得" + lstAccount.Count + "筆資料";
            return PartialView(lstAccount);
        }

        public ActionResult PlanAccountOfForm(string formid)
        {
            logger.Info("get plan account by form id=" + formid);
            List<PlanAccountFunction> lstAccount = service.getPlanAccount(null, null, null, "R", formid, null, null);
            return View(lstAccount);
        }
        public string getPlanAccountItem(string itemid)
        {
            logger.Info("get plan account item by id=" + itemid);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(service.getPlanAccountItem(itemid));
            logger.Info("plan account item  info=" + itemJson);
            return itemJson;
        }

        /// <summary>
        /// PlanAccountItem 註記刪除
        /// </summary>
        /// <param name="itemid"></param>
        /// <returns></returns>
        public String delPlanAccountItem(string itemid)
        {
            string msg = "刪除成功!!";
            if (null != itemid && itemid != "")
            {
                logger.Info("del plan account item by id=" + itemid);
            }
            else{
                itemid = Request["plan_account_id"];
                string allValue = "";
                foreach (string key in Request.Form.Keys)
                {
                    string value = Request.Form[key];
                    allValue = allValue + "{" + key + ":" + value + " }";
                }
                logger.Info("plan account item  info=" + allValue);
            }

            logger.Info("del plan account item by id=" + itemid);
            int i = service.delPlanAccountItem(itemid);
            return msg + "(" + i + ")";
        }
        public String updatePlanAccountItem(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "更新成功!!";

            PLAN_ACCOUNT item = new PLAN_ACCOUNT();
            item.PROJECT_ID = form["project_id"];
            item.PLAN_ACCOUNT_ID = int.Parse(form["plan_account_id"]);
            item.CONTRACT_ID = form["contract_id"];
            item.ACCOUNT_FORM_ID = form["account_form_id"];
            item.PAYMENT_DATE = Convert.ToDateTime(form.Get("date"));
            try
            {
                item.AMOUNT_PAID = decimal.Parse(form["amount_paid"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.PLAN_ACCOUNT_ID + " not paid amount:" + ex.Message);
            }
            try
            {
                item.AMOUNT_PAYABLE = decimal.Parse(form["amount_payable"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.PLAN_ACCOUNT_ID + " not payable amount:" + ex.Message);
            }
            item.ACCOUNT_TYPE = form["type"];
            logger.Debug("account type = " + form["type"]);
            item.ISDEBIT = form["isdebit"];
            item.STATUS = int.Parse(form["unRecordedFlag"]);
            item.CREATE_ID = form["create_id"];
            item.CHECK_NO = form["check_no"];
            item.PAYEE = form["payee"];
            item.REMARK = form["remark"];
            SYS_USER loginUser = (SYS_USER)Session["user"];
            item.MODIFY_ID = loginUser.USER_ID;
            item.MODIFY_DATE = DateTime.Now;
            int i = 0;
            i = service.updatePlanAccountItem(item);
            if (i == 0) { msg = service.message; }
            return msg;
        }
        //公司預算數與實際數查詢功能
        public ActionResult OperationExpSummary()
        {
            return View();
        }
        //查詢公司預算/實際數
        public ActionResult SearchExpSummary()
        {
            List<ExpenseBudgetSummary> ExpenseSummary = null;
            List<ExpenseBudgetSummary> BudgetSummary = null;

            ExpenseBudgetSummary Amt = null;
            ExpenseBudgetSummary ExpAmt = null;
            ExpenseBudgetModel viewModel = new ExpenseBudgetModel();
            if (null != Request["budgetyear"])
            {
                //取得預算數
                BudgetSummary = service.getExpBudgetByYear(int.Parse(Request["budgetyear"]));
                //取得發生數、
                ExpenseSummary = service.getExpSummaryByYear(int.Parse(Request["budgetyear"]));

                ExpAmt = service.getTotalOperationExpAmount(int.Parse(Request["budgetyear"]));
                Amt = service.getTotalExpBudgetAmount(int.Parse(Request["budgetyear"]));

                viewModel.BudgetSummary = BudgetSummary;
                viewModel.ExpenseSummary = ExpenseSummary;

                TempData["TotalAmt"] = String.Format("{0:#,##0.#}", Amt.TOTAL_BUDGET);
                TempData["TotalExpAmt"] = String.Format("{0:#,##0.#}", ExpAmt.TOTAL_OPERATION_EXP);
                TempData["budgetYear"] = Request["budgetyear"];
                return View("OperationExpSummary", viewModel);
            }

            TempData["budgetYear"] = Request["budgetyear"];
            return View("OperationExpSummary");
        }

        /// <summary>
        /// 下載費用表
        /// </summary>
        public void downLoadExpenseForm()
        {
            string formid = Request["formid"];
            service.getEXPByExpId(formid);
            if (null != service.formEXP)
            {
                ExpenseFormToExcel poi = new ExpenseFormToExcel();
                //檔案位置
                string fileLocation = poi.exportExcel(service.formEXP, service.EXPItem, service.siteEXPItem, service.ExpAmt, service.EarlyCumAmt, service.SiteEarlyCumAmt);
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

        /// <summary>
        /// 公司費用預算執行彙整表
        /// </summary>
        public void downLoadOPExpenseSummary()
        {
            int budgetYear = int.Parse(Request["budgetyear"]);
            ExpBudgetSummaryToExcel poi = new ExpBudgetSummaryToExcel();
            //檔案位置
            string fileLocation = poi.exportExcel(budgetYear);
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
        public ActionResult SearchCashInFlow()
        {

            List<PlanAccountFunction> CashInFlow = null;
            List<LoanTranactionFunction> LoanInFlow = null;
            CashFlowModel viewModel = new CashFlowModel();
            string projectname = "";
            string account_type = "R";
            string type = "I";
            CashInFlow = service.getPlanAccount(Request["payment_date"], projectname, projectname, account_type, projectname, Request["during_start"], Request["during_end"]);
            LoanInFlow = service.getLoanTranaction(type, Request["payment_date"], Request["during_start"], Request["during_end"]);
            viewModel.planAccount = CashInFlow;
            viewModel.finLoanTranaction = LoanInFlow;
            if (null != Request["payment_date"] && Request["payment_date"] != "")
            {
                ViewBag.SearchTerm = "查詢日期為 : " + Request["payment_date"];
            }
            else
            {
                ViewBag.SearchTerm = "查詢區間為 : " + Request["during_start"] + "~" + Request["during_end"];
            }
            return View("CashInFlowItem", viewModel);
        }
        public ActionResult SearchCashOutFlow()
        {
            List<PlanAccountFunction> CashOutFlow = null;
            List<LoanTranactionFunction> LoanOutFlow = null;
            List<PlanAccountFunction> OutFlowBalance = null;
            List<ExpenseFormFunction> ExpenseOutFlow = null;
            List<ExpenseFormFunction> ExpenseBudget = null;
            CashFlowModel viewModel = new CashFlowModel();
            string projectname = "";
            string account_type = "P";
            string type = "O";
            CashOutFlow = service.getPlanAccount(Request["payment_date"], projectname, projectname, account_type, projectname, Request["during_start"], Request["during_end"]);
            LoanOutFlow = service.getLoanTranaction(type, Request["payment_date"], Request["during_start"], Request["during_end"]);
            OutFlowBalance = service.getOutFlowBalanceByDate(Request["payment_date"], Request["during_start"], Request["during_end"]);
            ExpenseOutFlow = service.getExpenseOutFlowByDate(Request["payment_date"], Request["during_start"], Request["during_end"]);
            ExpenseBudget = service.getExpenseBudgetByDate(Request["payment_date"], Request["during_start"], Request["during_end"]);
            viewModel.planAccount = CashOutFlow;
            viewModel.finLoanTranaction = LoanOutFlow;
            viewModel.outFlowBalance = OutFlowBalance;
            viewModel.outFlowExp = ExpenseOutFlow;
            viewModel.expBudget = ExpenseBudget;

            if (null != Request["payment_date"] && Request["payment_date"] != "")
            {
                ViewBag.SearchTerm = "查詢日期為 : " + Request["payment_date"];
            }
            else
            {
                ViewBag.SearchTerm = "查詢區間為 : " + Request["during_start"] + "~" + Request["during_end"];
            }
            return View("CashOutFlowItem", viewModel);
        }
        public ActionResult showBudgetStatus()
        {
            string paymentDate = Request["paymentdate"];
            List<Budget4CashFow> lst = service.getBudgetStatus(paymentDate);
            return View(lst);
        }
    }
}

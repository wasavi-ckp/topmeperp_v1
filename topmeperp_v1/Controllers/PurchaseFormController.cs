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
using System.Data;
using System.Globalization;
using System.Web.Script.Serialization;

namespace topmeperp.Controllers
{
    public class PurchaseFormController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(InquiryController));
        PurchaseFormService service = new PurchaseFormService();

        // GET: PurchaseForm
        [topmeperp.Filter.AuthFilter]
        public ActionResult Index()
        {
            List<ProjectList> lstProject = PlanService.SearchProjectByName("", "專案執行",null);
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            //畫面上權限管理控制
            //頁面上使用ViewBag 定義開關\@ViewBag.F10005
            //由Session 取得權限清單
            List<SYS_FUNCTION> lstFunctions = (List<SYS_FUNCTION>)Session["functions"];
            //開關預設關閉
            @ViewBag.F10005 = "disabled";
            //輪巡功能清單，若全線存在則將開關打開 @ViewBag.F10005 = "";
            foreach (SYS_FUNCTION f in lstFunctions)
            {
                if (f.FUNCTION_ID == "F10005")
                {
                    @ViewBag.F10005 = "";
                }
            }
            return View(lstProject);
        }

        public ActionResult FormIndex(string id)
        {
            log.Info("purchase form index : projectid=" + id);
            ViewBag.projectId = id;
            return View();
        }
        // POST : Search
        [HttpPost]
        public ActionResult FormIndex(FormCollection f)
        {

            log.Info("projectid=" + Request["projectid"] + ",textCode1=" + Request["textCode1"] + ",textCode2=" + Request["textCode2"]);
            //加入刪除註記 預設 "N" /
            List<topmeperp.Models.PlanItem4Map> lstProject =
                service.getPlanItem(Request["chkEx"], Request["projectid"], Request["textCode1"], Request["textCode2"], Request["textSystemMain"], Request["textSystemSub"], Request["formName"], Request["supplier"], "N", "Y");
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            ViewBag.projectId = Request["projectid"];
            return View("FormIndex", lstProject);
        }
        //Create Purchasing Form
        [HttpPost]
        public ActionResult Create(PLAN_SUP_INQUIRY qf)
        {
            //取得專案編號
            log.Info("Project Id:" + Request["prjId"]);
            //取得專案名稱
            log.Info("Project Name:" + Request["prjId"]);
            //取得使用者勾選品項ID
            log.Info("item_list:" + Request["chkItem"]);
            log.Info("emptyform:" + Request["emptyform"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            log.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                log.Info("item_list return No.:" + lstItemId[i]);
            }
            //建立空白詢價單
            log.Info("create new form template");
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            qf.FORM_NAME = Request["formName"];
            if (null != Request["isWage"])
            {
                qf.ISWAGE = "Y";
            }
            qf.PROJECT_ID = Request["prjId"];
            qf.CREATE_ID = u.USER_ID;
            qf.CREATE_DATE = DateTime.Now;
            qf.OWNER_NAME = uInfo.USER_NAME;
            qf.OWNER_EMAIL = uInfo.EMAIL;
            qf.OWNER_TEL = uInfo.TEL;
            qf.OWNER_FAX = uInfo.FAX;
            PLAN_SUP_INQUIRY_ITEM item = new PLAN_SUP_INQUIRY_ITEM();
            string fid = service.newPlanForm(qf, lstItemId);
            //產生採購詢價單實體檔案(是否需先註解掉，因為空白詢價單不用產生實體檔，
            //樣本轉廠商採購單時再產生即可)
            service.getInqueryForm(fid);
            PurchaseFormtoExcel poi = new PurchaseFormtoExcel();
            poi.exportExcel4po(service.formInquiry, service.formInquiryItem, false, false);
            if (Request["emptyform"] == "E")
            {
                return Redirect("FormMainPage?id=" + qf.PROJECT_ID);
            }
            else
            {
                return Redirect("SinglePrjForm?id=" + fid);
            }
            //return RedirectToAction("InquiryMainPage","Inquiry", qf.PROJECT_ID);
        }
        public ActionResult FormMainPage(string id)
        {
            log.Info("purchase form by projectID =" + id + ",status=" + Request["status"] + ",type=" + Request["type"] + ",formname=" + Request["formname"]);
            PurchaseFormModel formData = new PurchaseFormModel();
            if (null != id && id != "")
            {
                ViewBag.projectid = id;
                TND_PROJECT p = service.getProjectById(id);
                ViewBag.projectName = p.PROJECT_NAME;
                //formData.planTemplateForm = service.getFormTemplateByProject(id);
                formData.planFormFromSupplier = service.getFormByProject(id, Request["status"], Request["type"], Request["formname"]);
                //formData.planForm4All = service.getFormByProject(id, Request["status"], "A", Request["formname"]);
            }

            ViewBag.Status = "有效";
            return View(formData);
        }
        //採發作業-空白詢價單管理功能
        public ActionResult FormTemplateMgr(string id)
        {
            log.Info("purchase form by projectID =" + id + ",status=" + Request["status"]);
            if (null != id && id != "")
            {
                ViewBag.projectid = id;
                TND_PROJECT p = service.getProjectById(id);
                ViewBag.projectName = p.PROJECT_NAME;
                PlanService ps = new PlanService();
                var priId = ps.getBudgetById(id);
                ViewBag.budgetdata = priId;
                ViewBag.budgetmessage = "尚未編列預算";
                string status = "有效";
                if (null != Request["status"])
                {
                    status = Request["status"];
                }
                service.getInquiryWithBudget(p, status);
                if (p.WAGE_MULTIPLIER == null)
                {
                    ViewBag.wageunitprice = "2500";
                }
                else
                {
                    ViewBag.wageunitprice = p.WAGE_MULTIPLIER;
                }
            }
            return View(service.POFormData);
        }
        //顯示單一詢價單、報價單功能
        public ActionResult SinglePrjForm(string id)
        {
            log.Info("http get mehtod:" + id);
            if (null == id || id == "")
            {
                id = Request["id"];
            }
            PurchaseFormDetail singleForm = new PurchaseFormDetail();
            service.getInqueryForm(id);
            singleForm.planForm = service.formInquiry;
            singleForm.planFormItem = service.formInquiryItem;
            singleForm.prj = service.getProjectById(singleForm.planForm.PROJECT_ID);
            ViewBag.targetSupplier = service.getSupplierContractByFormId(id); //判斷詢價單是否已寫入PLAN_ITEM以供發包採購
            log.Debug("Project ID:" + singleForm.prj.PROJECT_ID);
            //設定更新標籤 for 空白詢價單
            if (null != Request["update"])
            {
                ViewBag.update = "Y";
            }
            else
            {
                ViewBag.update = "N";
            }
            return View(singleForm);

        }

        public String UpdateFormName(FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得空白詢價單名稱
            if (form.Get("inquiryformid") != null)
            {
                string[] formId = form.Get("inquiryformid").Split(',');
                string[] formName = form.Get("formname").Split(',');
                List<PLAN_SUP_INQUIRY> lstItem = new List<PLAN_SUP_INQUIRY>();
                for (int j = 0; j < formId.Count(); j++)
                {
                    PLAN_SUP_INQUIRY item = new PLAN_SUP_INQUIRY();
                    item.INQUIRY_FORM_ID = formId[j];
                    item.FORM_NAME = formName[j];
                    log.Debug("plan form id =" + item.INQUIRY_FORM_ID + "，form name =" + item.FORM_NAME);
                    lstItem.Add(item);
                }
                int i = service.addFormName(lstItem);
                if (i == 0)
                {
                    msg = service.message;
                }
                else
                {
                    msg = "更新空白詢價單名稱成功";
                }
                return msg;
            }
            else
            {
                return "無詢價單名稱需要輸入，所以無法更新詢價單名稱";
            }
        }

        public String UpdatePrjForm(FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得供應商詢價單資料
            PLAN_SUP_INQUIRY fm = new PLAN_SUP_INQUIRY();
            SYS_USER loginUser = (SYS_USER)Session["user"];
            fm.SUPPLIER_ID = form.Get("Supplier").Substring(7).Trim();
            fm.PROJECT_ID = form.Get("projectid").Trim();
            fm.DUEDATE = Convert.ToDateTime(form.Get("inputdateline"));
            fm.OWNER_NAME = form.Get("inputowner").Trim();
            fm.OWNER_TEL = form.Get("inputphone").Trim();
            fm.OWNER_FAX = form.Get("inputownerfax").Trim();
            fm.OWNER_EMAIL = form.Get("inputowneremail").Trim();
            fm.FORM_NAME = form.Get("formname").Trim();
            fm.CREATE_ID = loginUser.USER_ID;
            fm.CREATE_DATE = DateTime.Now;
            fm.ISWAGE = "N";
            if (null != form.Get("isWage"))
            {
                fm.ISWAGE = form.Get("isWage").Trim();
            }
            TND_SUPPLIER s = service.getSupplierInfo(form.Get("Supplier").Substring(0, 7).Trim());
            fm.CONTACT_NAME = s.CONTACT_NAME;
            fm.CONTACT_EMAIL = s.CONTACT_EMAIL;
            string[] lstItemId = form.Get("formitemid").Split(',');
            log.Info("select count:" + lstItemId.Count());
            var j = 0;
            for (j = 0; j < lstItemId.Count(); j++)
            {
                log.Info("item_list return No.:" + lstItemId[j]);
            }
            string fid = service.addSupplierForm(fm, lstItemId);
            string[] lstPlanItem = form.Get("plan_item_id").Split(',');
            string[] lstQty = form.Get("formQty").Split(',');
            string[] lstPrice = form.Get("formunitprice").Split(',');
            string[] lstRemark = form.Get("remark").Split(',');
            List<PLAN_SUP_INQUIRY_ITEM> lstItem = new List<PLAN_SUP_INQUIRY_ITEM>();
            for (int i = 0; i < lstItemId.Count(); i++)
            {
                PLAN_SUP_INQUIRY_ITEM item = new PLAN_SUP_INQUIRY_ITEM();
                item.PLAN_ITEM_ID = lstPlanItem[i];
                if (lstRemark[i].ToString() == "")
                {
                    item.ITEM_REMARK = null;
                }
                else
                {
                    item.ITEM_REMARK = lstRemark[i];
                }
                if (lstQty[i].ToString() == "")
                {
                    item.ITEM_QTY = null;
                }
                if (lstPrice[i].ToString() == "")
                {
                    item.ITEM_UNIT_PRICE = null;
                }
                else
                {
                    item.ITEM_UNIT_PRICE = decimal.Parse(lstPrice[i]);
                }
                log.Debug("Plan Item Id=" + item.PLAN_ITEM_ID + ", Price =" + item.ITEM_UNIT_PRICE);
                lstItem.Add(item);
            }
            int k = service.refreshSupplierFormItem(fid, lstItem);

            //service.getInqueryForm(fid);
            //PurchaseFormtoExcel poi = new PurchaseFormtoExcel();
            //poi.exportExcel4po(service.formInquiry, service.formInquiryItem, false, true);
            if (fid == "")
            {
                msg = service.message;
            }
            else
            {
                msg = "新增供應商採購詢價單成功";
            }

            log.Info("Request:FORM_NAME=" + form["formname"] + "SUPPLIER_NAME =" + form["Supplier"]);
            return msg;
        }
        //更新採購廠商詢價單資料-new
        public String RefreshPrjForm(string id, FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得供應商詢價單資料
            PLAN_SUP_INQUIRY fm = new PLAN_SUP_INQUIRY();
            SYS_USER loginUser = (SYS_USER)Session["user"];
            fm.PROJECT_ID = form.Get("projectid").Trim();
            fm.ISWAGE = "N";
            if (null != form.Get("isWage"))
            {
                fm.ISWAGE = form.Get("isWage");
            }
            if (form.Get("supplier").Trim() != "")
            {
                fm.SUPPLIER_ID = form.Get("supplier").Trim();
                fm.DUEDATE = Convert.ToDateTime(form.Get("inputdateline"));
                fm.CONTACT_NAME = form.Get("inputcontact").Trim();
                fm.CONTACT_EMAIL = form.Get("inputemail").Trim();
            }
            fm.OWNER_NAME = form.Get("inputowner").Trim();
            fm.OWNER_TEL = form.Get("inputphone").Trim();
            fm.OWNER_FAX = form.Get("inputownerfax").Trim();
            fm.OWNER_EMAIL = form.Get("inputowneremail").Trim();
            fm.INQUIRY_FORM_ID = form.Get("inputformnumber").Trim();
            fm.FORM_NAME = form.Get("formname").Trim();
            fm.CREATE_ID = form.Get("createid").Trim();
            try
            {
                fm.CREATE_DATE = DateTime.ParseExact(form.Get("createdate"), "yyyy/MM/dd", CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                log.Error(ex.StackTrace);
            }
            fm.MODIFY_ID = loginUser.USER_ID;
            fm.MODIFY_DATE = DateTime.Now;
            string formid = form.Get("inputformnumber").Trim();

            string[] lstItemId = form.Get("formitemid").Split(',');
            string[] lstfItemId = form.Get("fItemId").Split(',');
            string[] lstfItemDesc = (string[])form.GetValue("fItemDesc").RawValue;
            string[] lstItemUnit = form.Get("fItemUnit").Split(',');
            string[] lstQty = form.Get("formQty").Split(',');
            string[] lstPrice = form.Get("formunitprice").Split(',');
            string[] lstRemark = form.Get("remark").Split(',');
            string[] lstPlanItemId = form.Get("plan_item_id").Split(',');
            List<PLAN_SUP_INQUIRY_ITEM> lstItem = new List<PLAN_SUP_INQUIRY_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_SUP_INQUIRY_ITEM item = new PLAN_SUP_INQUIRY_ITEM();
                item.INQUIRY_ITEM_ID = int.Parse(lstItemId[j]);
                item.ITEM_ID = lstfItemId[j];//ITEM_ID
                item.ITEM_DESC = lstfItemDesc[j];//ITEM_DESC
                item.ITEM_UNIT = lstItemUnit[j];//ITEM_UNIT
                item.PLAN_ITEM_ID = lstPlanItemId[j];//PLAN_ITEM_ID
                if (lstRemark[j].ToString() == "")
                {
                    item.ITEM_REMARK = null;
                }
                else
                {
                    item.ITEM_REMARK = lstRemark[j];
                }
                if (lstQty[j].ToString() == "")
                {
                    item.ITEM_QTY = null;
                }
                else
                {
                    item.ITEM_QTY = decimal.Parse(lstQty[j]);
                }
                if (lstPrice[j].ToString() == "")
                {
                    item.ITEM_UNIT_PRICE = null;
                }
                else
                {
                    item.ITEM_UNIT_PRICE = decimal.Parse(lstPrice[j]);
                }
                log.Debug("Item No=" + item.INQUIRY_ITEM_ID + ", Price =" + item.ITEM_UNIT_PRICE);
                lstItem.Add(item);
            }
            int i = service.refreshPlanSupplierForm(formid, fm, lstItem);
            if (service.message != "")
            {
                msg = service.message;
            }
            else
            {
                msg = "更新供應商採購詢價單成功，INQUIRY_FORM_ID =" + formid;
            }

            log.Info("Request: INQUIRY_FORM_ID = " + formid + "FORM_NAME =" + form["formname"] + "SUPPLIER_NAME =" + form["supplier"]);
            return msg;
        }

        //上傳廠商報價單
        public string FileUpload(HttpPostedFileBase file)
        {
            log.Info("Upload purchase form from supplier:" + Request["projectid"]);
            string projectid = Request["projectid"];
            string iswage = "N";
            string warringMsg = "";
            if (null != Request["isWage"])
            {
                log.Debug("isWage:" + Request["isWage"]);
                iswage = "Y";
            }
            //上傳至廠商報價單目錄
            if (null != file && file.ContentLength != 0)
            {
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + projectid + "/" + ContextService.quotesFolder, fileName);
                file.SaveAs(path);
                log.Info("Parser Excel File Begin:" + file.FileName);
                PurchaseFormtoExcel quoteFormService = new PurchaseFormtoExcel();
                try
                {
                    quoteFormService.convertInquiry2Plan(path, projectid, iswage);
                }
                catch (Exception ex)
                {
                    log.Error(ex.StackTrace);
                }
                int i = 0;
                //如果詢價單編號為空白，新增詢價單資料，否則更新相關詢價單資料-new
                log.Debug("Parser Excel File Finish!");

                if (service.getSupplierContractByFormId(quoteFormService.form.INQUIRY_FORM_ID) != false)
                {
                    warringMsg = "(此詢價單編號已被發包採購使用，資料已強制更新)";
                }
                if (null != quoteFormService.form.INQUIRY_FORM_ID && quoteFormService.form.INQUIRY_FORM_ID != "")
                {
                    log.Info("Update Plan Form for Inquiry:" + quoteFormService.form.INQUIRY_FORM_ID);
                    i = service.refreshPlanSupplierForm(quoteFormService.form.INQUIRY_FORM_ID, quoteFormService.form, quoteFormService.formItems);
                }
                else
                {
                    log.Info("Create New Plan Form for Inquiry:");
                    i = service.createPlanFormFromSupplier(quoteFormService.form, quoteFormService.formItems);
                }
                log.Info("add plan supplier form record count=" + i);
            }
            return "檔案匯入成功!!" + warringMsg;
        }
        //含工帶料報價單
        public string UploadInquiryAll()
        {
            log.Info("Upload purchase form from supplier:" + Request["projectid"]);
            string projectid = Request["projectid"];
            var file1 = Request.Files["file1"];
            //上傳至廠商報價單目錄
            if (null != file1 && file1.ContentLength != 0)
            {
                var fileName = Path.GetFileName(file1.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + projectid + "/" + ContextService.quotesFolder, fileName);
                file1.SaveAs(path);
                log.Info("Parser Excel File Begin:" + file1.FileName);
                PurchaseFormtoExcel quoteFormService = new PurchaseFormtoExcel();
                try
                {
                    quoteFormService.convertInquiryAll(path, projectid);
                }
                catch (Exception ex)
                {
                    log.Error(ex.StackTrace);
                    return ex.Message;
                }
                int i = 0;
                //如果詢價單編號為空白，新增詢價單資料，否則更新相關詢價單資料-new
                log.Debug("Parser Excel File Finish!");
                if (null != quoteFormService.form.INQUIRY_FORM_ID && quoteFormService.form.INQUIRY_FORM_ID != "")
                {
                    log.Info("Update Plan Form for Inquiry:" + quoteFormService.form.INQUIRY_FORM_ID);
                    i = service.refreshPlanSupplierForm(quoteFormService.form.INQUIRY_FORM_ID, quoteFormService.form, quoteFormService.formItems);
                }
                else
                {
                    log.Info("Create New Plan Form for Inquiry:");
                    i = service.createPlanFormFromSupplier(quoteFormService.form, quoteFormService.formItems);
                }
                log.Info("add plan supplier form record count=" + i);
            }
            return "檔案匯入成功!!";
        }
        //批次產生採購空白詢價單
        public string createPlanEmptyForm()
        {
            log.Info("project id=" + Request["projectid"]);
            SYS_USER u = (SYS_USER)Session["user"];
            int i = service.createPlanEmptyForm(Request["projectid"], u);
            return "共產生 " + i + "空白詢價單樣本!!";
        }
        /// <summary>
        /// 下載空白詢價單
        /// </summary>
        public void downLoadInquiryForm()
        {
            string formid = Request["formid"];
            bool isTemp = false;
            bool isReal = false;
            service.getInqueryForm(formid);
            if (null != Request["isTemp"] && Request["isTemp"] == "Y")
            {
                isTemp = true;
            }
            if (null != Request["isReal"] && Request["isReal"] == "Y")
            {
                isReal = true;
            }
            if (null != service.formInquiry)
            {
                PurchaseFormtoExcel poi = new PurchaseFormtoExcel();
                //檔案位置
                string fileLocation = poi.exportExcel4po(service.formInquiry, service.formInquiryItem, isTemp, isReal);
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
        /// 下載空白詢價單(含工帶料)
        /// </summary>
        public void downLoadInquiryForm4All()
        {
            string formid = Request["formid"];
            service.getInqueryForm(formid);
            if (null != service.formInquiry)
            {
                PurchaseFormtoExcel poi = new PurchaseFormtoExcel();
                //檔案位置
                string fileLocation = null;
                if (null == Request["isReal"])
                {
                    fileLocation = poi.exportExcel4poAll(service.formInquiry, service.formInquiryItem, false, false);
                }
                else
                {
                    fileLocation = poi.exportExcel4poAll(service.formInquiry, service.formInquiryItem, false, true);

                }
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
        //下載所有空白詢價單(採用zip 壓縮)
        public void downloadAllTemplate()
        {
            string projectid = Request["projectid"];
            log.Debug("create all template file by projectid=" + projectid);
            string zipFile = service.zipAllTemplate4Download(projectid);
            if (zipFile != "")
            {
                // 檔案名稱 HttpUtility.UrlEncode預設會以UTF8的編碼系統進行QP(Quoted - Printable)編碼，可以直接顯示的7 Bit字元(ASCII)就不用特別轉換。
                string filename = HttpUtility.UrlEncode(Path.GetFileName(zipFile));
                Response.Clear();
                Response.Charset = "utf-8";
                Response.ContentType = "text/zip";
                Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", filename));
                Response.WriteFile(zipFile);
                Response.End();
            }
        }
        //議約採購功能主頁
        public ActionResult PurchaseMain(string id, string formname, string iswage)
        {
            //傳入專案編號，
            log.Info("start project id=" + id);
            //取得專案基本資料
            ViewBag.id = id;
            PurchaseFormService service = new PurchaseFormService();
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            List<purchasesummary> lstforms = service.getPurchaseForm4Offer(id, formname, iswage);
            BudgetDataService bs = new BudgetDataService();
            DirectCost iteminfo = bs.getItemBudget(id, Request["textCode1"], Request["textCode2"], Request["textSystemMain"], Request["textSystemSub"], formname);
            ViewBag.itembudget = iteminfo.ITEM_BUDGET;
            ViewBag.itemwagebudget = iteminfo.ITEM_BUDGET_WAGE;
            ViewBag.SearchResult = "共取得" + lstforms.Count + "筆資料";
            return View(lstforms);
        }

        //採購比價功能資料頁
        public ActionResult Comparison(string id)
        {
            //傳入專案編號，
            log.Info("start project id=" + id);

            //取得專案基本資料
            PurchaseFormService service = new PurchaseFormService();
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.id = p.PROJECT_ID;
            ViewBag.projectName = p.PROJECT_NAME;
            SelectListItem empty = new SelectListItem();
            empty.Value = "";
            empty.Text = "";
            //取得主系統資料
            List<SelectListItem> selectMain = new List<SelectListItem>();
            foreach (string itm in service.getSystemMain(id))
            {
                log.Debug("Main System=" + itm);
                SelectListItem selectI = new SelectListItem();
                selectI.Value = itm;
                selectI.Text = itm;
                if (null != itm && "" != itm)
                {
                    selectMain.Add(selectI);
                }
            }
            // selectMain.Add(empty);
            ViewBag.SystemMain = selectMain;
            //取得次系統資料
            List<SelectListItem> selectSub = new List<SelectListItem>();
            foreach (string itm in service.getSystemSub(id))
            {
                log.Debug("Sub System=" + itm);
                SelectListItem selectI = new SelectListItem();
                selectI.Value = itm;
                selectI.Text = itm;
                if (null != itm && "" != itm)
                {
                    selectSub.Add(selectI);
                }
            }
            //selectSub.Add(empty);
            ViewBag.SystemSub = selectSub;
            string formName = Request["formName"];
            log.Debug("formName=" + formName);
            ViewBag.formName = formName;
            string iswage = Request["iswage"];
            if (iswage == "Y")
            {
                ViewBag.isWage = "checked";
            }

            //設定查詢條件
            return View();
        }

        //取得比價資料
        [HttpPost]
        public ActionResult ComparisonData(FormCollection form)
        {
            //傳入查詢條件
            log.Info("start project id=" + Request["id"] + ",TypeCode1=" + Request["typeCode1"] + ",typecode2=" + Request["typeCode2"] + ",SystemMain=" + Request["SystemMain"] + ",Sytem Sub=" + Request["SystemSub"] + ",Form Name=" + Request["formName"]);
            string iswage = "N";
            ViewBag.isWage = Request["isWage"];
            //取得備標品項與詢價資料
            try
            {
                if (null != Request["formName"] && "" != Request["formName"])
                {
                    if (null != Request["isWage"])
                    {
                        iswage = Request["isWage"];
                    }
                    DataTable dt = service.getComparisonDataToPivot(Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"], Request["formName"], iswage);
                    @ViewBag.ResultMsg = "共" + dt.Rows.Count + "筆";

                    //畫面上權限管理控制
                    //頁面上使用ViewBag 定義開關\@ViewBag.F10006
                    //由Session 取得權限清單
                    List<SYS_FUNCTION> lstFunctions = (List<SYS_FUNCTION>)Session["functions"];
                    //開關預設關閉
                    @ViewBag.F10006 = "hidden='hidden'";
                    //輪巡功能清單，若全線存在則將開關打開 @ViewBag.F10006 = "";
                    foreach (SYS_FUNCTION f in lstFunctions)
                    {
                        if (f.FUNCTION_ID == "F10006")
                        {
                            @ViewBag.F10006 = "";
                        }
                    }
                    string htmlString = "<table class='table table-bordered'><tr>";
                    //處理表頭
                    for (int i = 1; i < 7; i++)
                    {
                        log.Debug("column name=" + dt.Columns[i].ColumnName);
                        htmlString = htmlString + "<th>" + dt.Columns[i].ColumnName + "</th>";
                    }
                    //處理供應商表頭
                    Dictionary<string, COMPARASION_DATA_4PLAN> dirSupplierQuo = service.dirSupplierQuo;
                    log.Debug("Column Count=" + dt.Columns.Count);
                    List<string> list = new List<string>();
                    for (int i = 8; i < dt.Columns.Count; i++)
                    {
                        log.Debug("column name=" + dt.Columns[i].ColumnName);
                        string[] tmpString = dt.Columns[i].ColumnName.Split('|');
                        //<a href="/PurchaseForm/SinglePrjForm/@item.INQUIRY_FORM_ID" target="_blank">@item.INQUIRY_FORM_ID</a>
                        decimal tAmount = (decimal)dirSupplierQuo[tmpString[1]].TAmount;
                        string strAmout = string.Format("{0:C0}", tAmount);

                        htmlString = htmlString + "<th><table><tr><td>" + tmpString[0] + '(' + tmpString[2] + ')' +
                            "</td><button type='button' class='btn-xs' " + @ViewBag.F10006 + " onclick=\"clickSupplier('" + tmpString[1] + "','" + iswage + "')\"><span class='glyphicon glyphicon-ok' aria-hidden='true'></span></button>" +
                            "</button>" + "<button type = 'button' class='btn-xs' onclick=\"removeSupplier('" + tmpString[1] + "')\"><span class='glyphicon glyphicon-remove' aria-hidden='true'></span></button>" +
                            "<button type='button' class='btn-xs'><a href='/PurchaseForm/SinglePrjForm/" + tmpString[1] + "'" + " target='_blank'><span class='glyphicon glyphicon-list-alt' aria-hidden='true'></span></a>" +
                            "<br/><tr><td style = 'text-align:center;background-color:yellow;' > " + strAmout + "</td>" +
                            "</tr></table></th>";
                    }
                    htmlString = htmlString + "</tr>";
                    //處理資料表
                    foreach (DataRow dr in dt.Rows)
                    {
                        htmlString = htmlString + "<tr>";
                        for (int i = 1; i < 6; i++)
                        {
                            htmlString = htmlString + "<td>" + dr[i] + "</td>";
                        }
                        //單價欄位  <input type='text' id='cost_@item.INQUIRY_ITEM_ID' name='cost_@item.INQUIRY_ITEM_ID' size='5' />
                        //decimal price = decimal.Parse(dr[5].ToString());
                        if (dr[6].ToString() != "")
                        {
                            log.Debug("data row col 6=" + (decimal)dr[6]);
                            htmlString = htmlString + "<td><input type='text' id='cost_" + dr[1] + "' name='cost_" + dr[1] + "' size='5' value='" + String.Format("{0:N0}", (decimal)dr[6]) + "' /></td>";
                        }
                        else
                        {
                            htmlString = htmlString + "<td></td>";
                        }
                        //String.Format("{0:C}", 0);
                        //處理報價資料
                        for (int i = 8; i < dt.Columns.Count; i++)
                        {
                            //<td><button class="btn-link" onclick="clickPrice('@item.INQUIRY_ITEM_ID', '@item.QUOTATION_PRICE')">@item.QUOTATION_PRICE</button> </td>
                            if (dr[i].ToString() != "")
                            {
                                htmlString = htmlString + "<td><button class='btn-link' onclick=\"clickPrice('" + dr[1] + "', '" + dr[i] + "','" + iswage + "')\">" + String.Format("{0:N0}", (decimal)dr[i]) + "</button> </td>";
                            }
                            else
                            {
                                htmlString = htmlString + "<td></td>";
                            }
                        }
                        htmlString = htmlString + "</tr>";
                    }
                    htmlString = htmlString + "</table>";
                    getBudget(Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"], Request["formName"]);
                    //產生畫面
                    IHtmlString str = new HtmlString(htmlString);
                    ViewBag.htmlString = str;
                }
                else
                {
                    if (null != Request["isWage"])
                    {
                        iswage = Request["isWage"];
                    }
                    DataTable dt = service.getComparisonDataToPivot(Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"], Request["formName"], iswage);
                    @ViewBag.ResultMsg = "共" + dt.Rows.Count + "筆";

                    //畫面上權限管理控制
                    //頁面上使用ViewBag 定義開關\@ViewBag.F10006
                    //由Session 取得權限清單
                    List<SYS_FUNCTION> lstFunctions = (List<SYS_FUNCTION>)Session["functions"];
                    //開關預設關閉
                    @ViewBag.F10006 = "disabled";
                    //輪巡功能清單，若全線存在則將開關打開 @ViewBag.F10006 = "";
                    foreach (SYS_FUNCTION f in lstFunctions)
                    {
                        if (f.FUNCTION_ID == "F10006")
                        {
                            @ViewBag.F10006 = "";
                        }
                    }
                    string htmlString = "<table class='table table-bordered'><tr>";
                    //處理表頭
                    for (int i = 1; i < 7; i++)
                    {
                        log.Debug("column name=" + dt.Columns[i].ColumnName);
                        htmlString = htmlString + "<th>" + dt.Columns[i].ColumnName + "</th>";
                    }
                    //處理供應商表頭
                    Dictionary<string, COMPARASION_DATA_4PLAN> dirSupplierQuo = service.dirSupplierQuo;
                    log.Debug("Column Count=" + dt.Columns.Count);
                    List<string> list = new List<string>();
                    for (int i = 7; i < dt.Columns.Count; i++)
                    {
                        log.Debug("column name=" + dt.Columns[i].ColumnName);
                        string[] tmpString = dt.Columns[i].ColumnName.Split('|');
                        //<a href="/PurchaseForm/SinglePrjForm/@item.INQUIRY_FORM_ID" target="_blank">@item.INQUIRY_FORM_ID</a>
                        decimal tAmount = 0;
                        string strAmout = "0";
                        try
                        {
                            tAmount = (decimal)dirSupplierQuo[tmpString[1]].TAmount;
                            strAmout = string.Format("{0:C0}", tAmount);
                        }
                        catch (Exception ex)
                        {
                            log.Error(dt.Columns[i].ColumnName + " not have Budget" + ex.Message);
                            log.Debug(ex.StackTrace);
                        }

                        htmlString = htmlString + "<th><table><tr><td>" + tmpString[0] + '(' + tmpString[2] + ')' +
                            "</td><button type='button' class='btn-xs' " + @ViewBag.F10006 + " onclick=\"clickSupplier('" + tmpString[1] + "','" + iswage + "')\"><span class='glyphicon glyphicon-ok' aria-hidden='true'></span></button>" +
                            "</button>" + "<button type = 'button' class='btn-xs' onclick=\"removeSupplier('" + tmpString[1] + "')\"><span class='glyphicon glyphicon-remove' aria-hidden='true'></span></button>" +
                            "<button type='button' class='btn-xs'><a href='/PurchaseForm/SinglePrjForm/" + tmpString[1] + "'" + " target='_blank'><span class='glyphicon glyphicon-list-alt' aria-hidden='true'></span></a>" +
                            "<br/><tr><td style = 'text-align:center;background-color:yellow;' > " + strAmout + "</td>" +
                            "</tr></table></th>";
                    }
                    htmlString = htmlString + "</tr>";
                    //處理資料表
                    foreach (DataRow dr in dt.Rows)
                    {
                        htmlString = htmlString + "<tr>";
                        for (int i = 1; i < 6; i++)
                        {
                            htmlString = htmlString + "<td>" + dr[i] + "</td>";
                        }
                        //單價欄位  <input type='text' id='cost_@item.INQUIRY_ITEM_ID' name='cost_@item.INQUIRY_ITEM_ID' size='5' />
                        //decimal price = decimal.Parse(dr[5].ToString());
                        if (dr[6].ToString() != "")
                        {
                            log.Debug("data row col 6=" + (decimal)dr[6]);
                            htmlString = htmlString + "<td><input type='text' id='cost_" + dr[1] + "' name='cost_" + dr[1] + "' size='5' value='" + String.Format("{0:N2}", (decimal)dr[6]) + "' /></td>";
                        }
                        else
                        {
                            htmlString = htmlString + "<td></td>";
                        }
                        //String.Format("{0:C}", 0);
                        //處理報價資料
                        for (int i = 7; i < dt.Columns.Count; i++)
                        {
                            //<td><button class="btn-link" onclick="clickPrice('@item.INQUIRY_ITEM_ID', '@item.QUOTATION_PRICE')">@item.QUOTATION_PRICE</button> </td>
                            if (dr[i].ToString() != "")
                            {
                                htmlString = htmlString + "<td><button class='btn-link' onclick=\"clickPrice('" + dr[1] + "', '" + dr[i] + "','" + iswage + "')\">" + String.Format("{0:N2}", (decimal)dr[i]) + "</button> </td>";
                            }
                            else
                            {
                                htmlString = htmlString + "<td></td>";
                            }
                        }
                        htmlString = htmlString + "</tr>";
                    }
                    htmlString = htmlString + "</table>";
                    //加入預算資料
                    getBudget(Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"], Request["formName"]);
                    //產生畫面
                    IHtmlString str = new HtmlString(htmlString);
                    ViewBag.htmlString = str;
                }
            }
            catch (Exception e)
            {
                log.Error("Ex" + e.Message + "," + e.StackTrace);
                ViewBag.htmlString = e.Message;
            }
            return PartialView();
        }

        private void getBudget(string projectid, string typecode1, string typecode2, string systemmain, string systemsub, string formname)
        {
            //加入預算資料
            BudgetDataService bs = new BudgetDataService();
            DirectCost iteminfo = bs.getItemBudget(projectid, typecode1, typecode2, systemmain, systemsub, formname);
            ViewBag.itembudget = iteminfo.ITEM_BUDGET;
            ViewBag.itemwagebudget = iteminfo.ITEM_BUDGET_WAGE;
        }

        //工資比價功能資料頁
        public ActionResult Comparison4Wage(string id)
        {
            //傳入專案編號，
            log.Info("start project id=" + id);

            //取得專案基本資料
            PurchaseFormService service = new PurchaseFormService();
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.id = p.PROJECT_ID;
            ViewBag.projectName = p.PROJECT_NAME;
            SelectListItem empty = new SelectListItem();
            empty.Value = "";
            empty.Text = "";
            //取得主系統資料
            List<SelectListItem> selectMain = new List<SelectListItem>();
            foreach (string itm in service.getSystemMain(id))
            {
                log.Debug("Main System=" + itm);
                SelectListItem selectI = new SelectListItem();
                selectI.Value = itm;
                selectI.Text = itm;
                if (null != itm && "" != itm)
                {
                    selectMain.Add(selectI);
                }
            }
            // selectMain.Add(empty);
            ViewBag.SystemMain = selectMain;
            //取得次系統資料
            List<SelectListItem> selectSub = new List<SelectListItem>();
            foreach (string itm in service.getSystemSub(id))
            {
                log.Debug("Sub System=" + itm);
                SelectListItem selectI = new SelectListItem();
                selectI.Value = itm;
                selectI.Text = itm;
                if (null != itm && "" != itm)
                {
                    selectSub.Add(selectI);
                }
            }
            //selectSub.Add(empty);
            ViewBag.SystemSub = selectSub;
            //設定查詢條件
            return View();
        }

        //取得工資比價資料
        [HttpPost]
        public ActionResult ComparisonData4Wage(FormCollection form)
        {
            //傳入查詢條件
            log.Info("start project id=" + Request["id"] + ",TypeCode1=" + Request["typeCode1"] + ",typecode2=" + Request["typeCode2"] + ",SystemMain=" + Request["SystemMain"] + ",Sytem Sub=" + Request["SystemSub"] + ",Form Name=" + Request["formName"]);
            string iswage = "Y";
            //取得備標品項與詢價資料
            try
            {
                if (null != Request["formName"] && "" != Request["formName"])
                {
                    DataTable dt = service.getComparisonDataToPivot(Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"], Request["formName"], iswage);
                    @ViewBag.ResultMsg = "共" + dt.Rows.Count + "筆";

                    //畫面上權限管理控制
                    //頁面上使用ViewBag 定義開關\@ViewBag.F10006
                    //由Session 取得權限清單
                    List<SYS_FUNCTION> lstFunctions = (List<SYS_FUNCTION>)Session["functions"];
                    //開關預設關閉
                    @ViewBag.F10006 = "disabled";
                    //輪巡功能清單，若全線存在則將開關打開 @ViewBag.F10006 = "";
                    foreach (SYS_FUNCTION f in lstFunctions)
                    {
                        if (f.FUNCTION_ID == "F10006")
                        {
                            @ViewBag.F10006 = "";
                        }
                    }
                    string htmlString = "<table class='table table-bordered'><tr>";
                    //處理表頭
                    for (int i = 1; i < 7; i++)
                    {
                        log.Debug("column name=" + dt.Columns[i].ColumnName);
                        htmlString = htmlString + "<th>" + dt.Columns[i].ColumnName + "</th>";
                    }
                    //處理供應商表頭
                    Dictionary<string, COMPARASION_DATA_4PLAN> dirSupplierQuo = service.dirSupplierQuo;
                    log.Debug("Column Count=" + dt.Columns.Count);
                    List<string> list = new List<string>();
                    for (int i = 8; i < dt.Columns.Count; i++)
                    {
                        log.Debug("column name=" + dt.Columns[i].ColumnName);
                        string[] tmpString = dt.Columns[i].ColumnName.Split('|');
                        //<a href="/PurchaseForm/SinglePrjForm/@item.INQUIRY_FORM_ID" target="_blank">@item.INQUIRY_FORM_ID</a>
                        decimal tAmount = (decimal)dirSupplierQuo[tmpString[1]].TAmount;
                        string strAmout = string.Format("{0:C0}", tAmount);
                        decimal avgPrice = (decimal)dirSupplierQuo[tmpString[1]].AvgMPrice;
                        string strAvgPrice = string.Format("{0:N0}", avgPrice);

                        htmlString = htmlString + "<th><table><tr><td>" + tmpString[0] + '(' + tmpString[2] + ')' +
                            "</td><button type='button' class='btn-xs' " + @ViewBag.F10006 + " onclick=\"clickSupplier('" + tmpString[1] + "','" + iswage + "')\"><span class='glyphicon glyphicon-ok' aria-hidden='true'></span></button>" +
                            "</button>" + "<button type = 'button' class='btn-xs' onclick=\"removeSupplier('" + tmpString[1] + "')\"><span class='glyphicon glyphicon-remove' aria-hidden='true'></span></button>" +
                            "<button type='button' class='btn-xs'><a href='/PurchaseForm/SinglePrjForm/" + tmpString[1] + "'" + " target='_blank'><span class='glyphicon glyphicon-list-alt' aria-hidden='true'></span></a>" +
                            "<br/><tr><td style = 'text-align:center;background-color:pink;'>工資/天 :" + strAvgPrice + "</td>" +
                            "<td style = 'text-align:center;background-color:yellow;' >總價 :" + strAmout + "</td>" +
                            "</tr></table></th>";
                    }
                    htmlString = htmlString + "</tr>";
                    //處理資料表
                    foreach (DataRow dr in dt.Rows)
                    {
                        htmlString = htmlString + "<tr>";
                        for (int i = 1; i < 6; i++)
                        {
                            htmlString = htmlString + "<td>" + dr[i] + "</td>";
                        }
                        //單價欄位  <input type='text' id='cost_@item.INQUIRY_ITEM_ID' name='cost_@item.INQUIRY_ITEM_ID' size='5' />
                        //decimal price = decimal.Parse(dr[5].ToString());
                        if (dr[6].ToString() != "")
                        {
                            log.Debug("data row col 6=" + (decimal)dr[6]);
                            htmlString = htmlString + "<td><input type='text' id='cost_" + dr[1] + "' name='cost_" + dr[1] + "' size='5' value='" + String.Format("{0:N0}", (decimal)dr[6]) + "' /></td>";
                        }
                        else
                        {
                            htmlString = htmlString + "<td></td>";
                        }
                        //String.Format("{0:C}", 0);
                        //處理報價資料
                        for (int i = 8; i < dt.Columns.Count; i++)
                        {
                            //<td><button class="btn-link" onclick="clickPrice('@item.INQUIRY_ITEM_ID', '@item.QUOTATION_PRICE')">@item.QUOTATION_PRICE</button> </td>
                            if (dr[i].ToString() != "")
                            {
                                htmlString = htmlString + "<td><button class='btn-link' onclick=\"clickPrice('" + dr[1] + "', '" + dr[i] + "','" + iswage + "')\">" + String.Format("{0:N0}", (decimal)dr[i]) + "</button> </td>";
                            }
                            else
                            {
                                htmlString = htmlString + "<td></td>";
                            }
                        }
                        htmlString = htmlString + "</tr>";
                    }
                    htmlString = htmlString + "</table>";
                    //加入預算資料
                    BudgetDataService bs = new BudgetDataService();
                    DirectCost iteminfo = bs.getItemBudget(Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"], Request["formName"]);
                    ViewBag.itembudget = iteminfo.ITEM_BUDGET;
                    ViewBag.itemwagebudget = iteminfo.ITEM_BUDGET_WAGE;
                    //產生畫面
                    IHtmlString str = new HtmlString(htmlString);
                    ViewBag.htmlString = str;
                }
                else
                {
                    DataTable dt = service.getComparisonDataToPivot(Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"], Request["formName"], iswage);
                    @ViewBag.ResultMsg = "共" + dt.Rows.Count + "筆";

                    //畫面上權限管理控制
                    //頁面上使用ViewBag 定義開關\@ViewBag.F10006
                    //由Session 取得權限清單
                    List<SYS_FUNCTION> lstFunctions = (List<SYS_FUNCTION>)Session["functions"];
                    //開關預設關閉
                    @ViewBag.F10006 = "disabled";
                    //輪巡功能清單，若全線存在則將開關打開 @ViewBag.F10006 = "";
                    foreach (SYS_FUNCTION f in lstFunctions)
                    {
                        if (f.FUNCTION_ID == "F10006")
                        {
                            @ViewBag.F10006 = "";
                        }
                    }
                    string htmlString = "<table class='table table-bordered'><tr>";
                    //處理表頭
                    for (int i = 1; i < 7; i++)
                    {
                        log.Debug("column name=" + dt.Columns[i].ColumnName);
                        htmlString = htmlString + "<th>" + dt.Columns[i].ColumnName + "</th>";
                    }
                    //處理供應商表頭
                    Dictionary<string, COMPARASION_DATA_4PLAN> dirSupplierQuo = service.dirSupplierQuo;
                    log.Debug("Column Count=" + dt.Columns.Count);
                    List<string> list = new List<string>();
                    for (int i = 7; i < dt.Columns.Count; i++)
                    {
                        log.Debug("column name=" + dt.Columns[i].ColumnName);
                        string[] tmpString = dt.Columns[i].ColumnName.Split('|');
                        //<a href="/PurchaseForm/SinglePrjForm/@item.INQUIRY_FORM_ID" target="_blank">@item.INQUIRY_FORM_ID</a>
                        decimal tAmount = (decimal)dirSupplierQuo[tmpString[1]].TAmount;
                        string strAmout = string.Format("{0:C0}", tAmount);
                        decimal avgPrice = (decimal)dirSupplierQuo[tmpString[1]].AvgMPrice;
                        string strAvgPrice = string.Format("{0:N0}", avgPrice);

                        htmlString = htmlString + "<th><table><tr><td>" + tmpString[0] + '(' + tmpString[2] + ')' +
                            "</td><button type='button' class='btn-xs' " + @ViewBag.F10006 + " onclick=\"clickSupplier('" + tmpString[1] + "','" + iswage + "')\"><span class='glyphicon glyphicon-ok' aria-hidden='true'></span></button>" +
                            "</button>" + "<button type = 'button' class='btn-xs' onclick=\"removeSupplier('" + tmpString[1] + "')\"><span class='glyphicon glyphicon-remove' aria-hidden='true'></span></button>" +
                            "<button type='button' class='btn-xs'><a href='/PurchaseForm/SinglePrjForm/" + tmpString[1] + "'" + " target='_blank'><span class='glyphicon glyphicon-list-alt' aria-hidden='true'></span></a>" +
                            "<br/><tr><td style = 'text-align:center;background-color:pink;'>工資/天 :" + strAvgPrice + "</td>" +
                            "<td style = 'text-align:center;background-color:yellow;'>總價 :" + strAmout + "</td>" +
                            "</tr></table></th>";
                    }
                    htmlString = htmlString + "</tr>";
                    //處理資料表
                    foreach (DataRow dr in dt.Rows)
                    {
                        htmlString = htmlString + "<tr>";
                        for (int i = 1; i < 6; i++)
                        {
                            htmlString = htmlString + "<td>" + dr[i] + "</td>";
                        }
                        //單價欄位  <input type='text' id='cost_@item.INQUIRY_ITEM_ID' name='cost_@item.INQUIRY_ITEM_ID' size='5' />
                        //decimal price = decimal.Parse(dr[5].ToString());
                        if (dr[6].ToString() != "")
                        {
                            log.Debug("data row col 6=" + (decimal)dr[6]);
                            htmlString = htmlString + "<td><input type='text' id='cost_" + dr[1] + "' name='cost_" + dr[1] + "' size='5' value='" + String.Format("{0:N0}", (decimal)dr[6]) + "' /></td>";
                        }
                        else
                        {
                            htmlString = htmlString + "<td></td>";
                        }
                        //String.Format("{0:C}", 0);
                        //處理報價資料
                        for (int i = 7; i < dt.Columns.Count; i++)
                        {
                            //<td><button class="btn-link" onclick="clickPrice('@item.INQUIRY_ITEM_ID', '@item.QUOTATION_PRICE')">@item.QUOTATION_PRICE</button> </td>
                            if (dr[i].ToString() != "")
                            {
                                htmlString = htmlString + "<td><button class='btn-link' onclick=\"clickPrice('" + dr[1] + "', '" + dr[i] + "','" + iswage + "')\">" + String.Format("{0:N0}", (decimal)dr[i]) + "</button> </td>";
                            }
                            else
                            {
                                htmlString = htmlString + "<td></td>";
                            }
                        }
                        htmlString = htmlString + "</tr>";
                    }
                    htmlString = htmlString + "</table>";
                    //加入預算資料
                    BudgetDataService bs = new BudgetDataService();
                    DirectCost iteminfo = bs.getItemBudget(Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"], Request["formName"]);
                    ViewBag.itembudget = iteminfo.ITEM_BUDGET;
                    ViewBag.itemwagebudget = iteminfo.ITEM_BUDGET_WAGE;
                    //產生畫面
                    IHtmlString str = new HtmlString(htmlString);
                    ViewBag.htmlString = str;
                }
            }
            catch (Exception e)
            {
                log.Error("Ex" + e.Message);
                ViewBag.htmlString = e.Message;
            }
            return PartialView();
        }

        //移除議價資料
        public string RemoveSupplierForm(string formid)
        {
            log.Info("formid=" + Request["formid"]);
            int i = service.removeSuplplierFormFromQuote(Request["formid"]);
            return "更新成功!!";
        }
        //更新單項成本資料
        public string UpdateCost4Item()
        {
            log.Info("PLanItemID=" + Request["pitmid"] + ",Cost=" + Request["price"] + ",iswage=" + Request["iswage"]);
            try
            {
                decimal cost = decimal.Parse(Request["price"]);
                service.updateCostFromQuote(Request["pitmid"], cost, Request["iswage"]);
            }
            catch (Exception ex)
            {
                log.Error(ex.StackTrace);
                return "更新失敗(請檢查資料格式是否有誤)!!";
            }
            return "更新成功!!";
        }
        //依據詢價單內容，更新得標標單品項所有單價
        public string BatchUpdateCost(string formid)
        {
            SYS_USER u = (SYS_USER)Session["user"];
            log.Info("project_id=" + Request["projectid"] + "formid=" + Request["formid"] + ",iswage=" + Request["iswage"]);
            int i = service.batchUpdateCostFromQuote(Request["formid"], Request["iswage"]);
            //增加項發包紀錄
            service.addContract4SupplierRecord(Request["projectid"], Request["formid"], u);
            return "更新成功!!";
        }
        /// <summary>
        /// 發包廠商之明細資料:
        /// 單一項目可發包多家廠商
        /// </summary>
        /// <returns></returns>
        public ActionResult Contract4SupplierIndex()
        {
            string projectid = Request["projectid"];
            PurchaseFormModel formData = new PurchaseFormModel();
            if (null != projectid && projectid != "")
            {
                ViewBag.projectid = projectid;
                TND_PROJECT p = service.getProjectById(projectid);
                ViewBag.projectName = p.PROJECT_NAME;
                ViewBag.formName = Request["formname"];
                string iswage = Request["iswage"] == "" ? "N" : "Y";
                formData.planFormFromSupplier = service.getContractForm(projectid, "有效", iswage, Request["formname"]);
            }
            getBudget(projectid, null, null, null, null, Request["formname"]);
            return View(formData);
        }
        /// <summary>
        /// 註銷發包廠商合約報價單資料
        /// </summary>
        /// <returns></returns>
        public string cancelContract()
        {
            SYS_USER u = (SYS_USER)Session["user"];
            string projectid = Request["projectid"];
            string formid = Request["formid"];
            int i = service.cancelContractForm(projectid, formid);
            log.Info(u.USER_ID + " cancel project" + projectid + "form Id=" + formid);
            return "廠商合約已取消("+ i +")";
        }
        public void changeStatus()
        {
            string formid = Request["formId"];
            string status = Request["status"];
            log.Debug("change form status:" + formid + ",status=" + status);
            service.changePlanFormStatus(formid, status);
        }
        //取得採購合約資料
        public ActionResult PurchasingContract(string id)
        {
            //傳入專案編號，
            log.Info("start project id=" + id);
            //取得專案基本資料
            ViewBag.id = id;
            PurchaseFormService service = new PurchaseFormService();
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            //取得採購合約金額與預算等資料
            List<plansummary> lstContract = null;
            List<plansummary> lstWageContract = null;
            ContractModels contract = new ContractModels();
            lstContract = service.getPlanContract(id);
            lstWageContract = service.getPlanContract4Wage(id);
            contract.contractItems = lstContract;
            contract.wagecontractItems = lstWageContract;
            plansummary contractAmount = service.getPlanContractAmount(id);
            ViewBag.budget = contractAmount.TOTAL_BUDGET;
            ViewBag.cost = contractAmount.TOTAL_COST;
            ViewBag.revenue = contractAmount.TOTAL_REVENUE;
            ViewBag.profit = contractAmount.TOTAL_PROFIT;
            ViewBag.SearchResult = "共取得" + lstContract.Count + "筆資料";
            ViewBag.Result = lstContract.Count;
            ViewBag.Result4Wage = lstWageContract.Count;
            return View(contract);
        }

        //取得合約付款條件
        public string getPaymentTerms(string contractid)
        {
            log.Info("access the terms of payment by:" + Request["contractid"]);
            PurchaseFormService service = new PurchaseFormService();
            if (contractid.Substring(0, 5) != "Owner")
            {
                service.getInqueryForm(contractid);
                PLAN_SUP_INQUIRY f = service.formInquiry;
                int i = service.addContractId(f.PROJECT_ID);
                int j = service.addContractIdForWage(f.PROJECT_ID);
            }
            else
            {
                contractid = contractid.Substring(5);
            }
            log.Info("contractid =" + contractid);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(service.getPaymentTerm(contractid, contractid));
            log.Info("plan payment terms info=" + itemJson);
            return itemJson;
        }

        //寫入合約付款條件
        public String addPaymentTerms(FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "更新成功!!";

            PLAN_PAYMENT_TERMS item = new PLAN_PAYMENT_TERMS();
            item.PROJECT_ID = form["project_id"];
            item.CONTRACT_ID = form["contract_id"];
            //付款方式與日期 : O:每月上下兩期,T:每月付款一次
            item.PAYMENT_FREQUENCY = form["payfrequency"];
            //付款日期定義 O 每月兩次: DATE_1、DATE_2
            try
            {
                item.DATE_1 = decimal.Parse(form["date1"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not date_1:" + ex.Message);
            }
            try
            {
                item.DATE_2 = decimal.Parse(form["date2"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not date_2:" + ex.Message);
            }
            //付款日期定義 T 每月一次: DATE_3
            try
            {
                item.DATE_3 = decimal.Parse(form["date3"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not date_3:" + ex.Message);
            }

            #region 按期估驗 付款方式 P
            //付款比例與票期 : P:按期估驗 S:階段付款 
            item.PAYMENT_TERMS = form["payterms"];

            //工資、設備、連工帶料 (for 按期估驗P使用)
            item.PAYMENT_TYPE = form["paymenttype"];
            //現金票 : ""、電匯、現金、支票
            item.PAYMENT_CASH = form["paymentcash"];

            //估驗階段第2階段 付款方式 支票、現金、電匯
            item.PAYMENT_UP_TO_U_1 = form["payment_1"];
            //估驗比率，第2部分 開票付款延遲天數
            try
            {
                item.PAYMENT_UP_TO_U_DATE1 = decimal.Parse(form["payment_date1"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not payment_date1:" + ex.Message);
            }
            //估驗階段第3階段 付款方式 支票、現金、電匯
            item.PAYMENT_UP_TO_U_2 = form["payment_2"];
            //估驗比率，第3部分  開票付款延遲天數
            try
            {
                item.PAYMENT_UP_TO_U_DATE2 = decimal.Parse(form["payment_date2"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not payment_date2:" + ex.Message);
            }

            //預付款 付款比例
            try
            {
                item.PAYMENT_ADVANCE_RATIO = decimal.Parse(form["paymentadvance"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not paymentadvance:" + ex.Message);
            }

            //預付款 現金付款比例
            try
            {
                item.PAYMENT_ADVANCE_CASH_RATIO = decimal.Parse(form["paymentadvance_cash"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not paymentadvance_cash:" + ex.Message);
            }
            //預付款 2 支票付款比例
            try
            {
                item.PAYMENT_ADVANCE_1_RATIO = decimal.Parse(form["paymentadvance_1"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not paymentadvance_1:" + ex.Message);
            }
            //預付款 3 支票付款比例
            try
            {
                item.PAYMENT_ADVANCE_2_RATIO = decimal.Parse(form["paymentadvance_2"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not paymentadvance_2:" + ex.Message);
            }
            //估驗款 付款比例
            try
            {
                item.PAYMENT_ESTIMATED_RATIO = decimal.Parse(form["paymentestimated"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not paymentestimated:" + ex.Message);
            }
            //估驗款 現金比例
            try
            {
                item.PAYMENT_ESTIMATED_CASH_RATIO = decimal.Parse(form["paymentestimated_cash"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not paymentestimated_cash:" + ex.Message);
            }
            //估驗款 2 支票金額比例
            try
            {
                item.PAYMENT_ESTIMATED_1_RATIO = decimal.Parse(form["paymentestimated_1"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not paymentestimated_1:" + ex.Message);
            }
            //估驗款 3 支票金額比例
            try
            {
                item.PAYMENT_ESTIMATED_2_RATIO = decimal.Parse(form["paymentestimated_2"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not paymentestimated_2:" + ex.Message);
            }

            //保留款 保留款比例
            try
            {
                item.PAYMENT_RETENTION_RATIO = decimal.Parse(form["paymentretention"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not paymentretention:" + ex.Message);
            }
            //保留款 現金比例
            try
            {
                item.PAYMENT_RETENTION_CASH_RATIO = decimal.Parse(form["paymentretention_cash"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not paymentretention_cash:" + ex.Message);
            }
            //保留款 2 支票金額比例
            try
            {
                item.PAYMENT_RETENTION_1_RATIO = decimal.Parse(form["paymentretention_1"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not paymentretention_1:" + ex.Message);
            }
            //保留款 3 支票金額比例
            try
            {
                item.PAYMENT_RETENTION_2_RATIO = decimal.Parse(form["paymentretention_2"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not paymentretention_2:" + ex.Message);
            }
            #endregion

            #region 階段付款 付款方式 S
            // 現金付款方式:""、電匯、現金、支票
            item.USANCE_CASH = form["usancecash"];
            // 支票付款1:電匯、現金、支票
            item.USANCE_UP_TO_U_1 = form["usance_1"];
            // 支票付款1 延遲天數 
            try
            {
                item.USANCE_UP_TO_U_DATE1 = decimal.Parse(form["usance_date1"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usance_date1:" + ex.Message);
            } 
            // 支票付款2 延遲天數 
            try
            {
                item.USANCE_UP_TO_U_DATE2 = decimal.Parse(form["usance_date2"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usance_date2:" + ex.Message);
            }

            // 支票付款2:電匯、現金、支票
            item.USANCE_UP_TO_U_2 = form["usance_2"];
            //訂金比例 (合約)
            try
            {
                item.USANCE_ADVANCE_RATIO = decimal.Parse(form["usanceadvance"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usanceadvance:" + ex.Message);
            }
            //訂金 現金比例
            try
            {
                item.USANCE_ADVANCE_CASH_RATIO = decimal.Parse(form["usanceadvance_cash"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usanceadvance_cash:" + ex.Message);
            }
            //訂金，支票 2 付款比例
            try
            {
                item.USANCE_ADVANCE_1_RATIO = decimal.Parse(form["usanceadvance_1"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usanceadvance_1:" + ex.Message);
            }
            //訂金，支票 3 付款比例
            try
            {
                item.USANCE_ADVANCE_2_RATIO = decimal.Parse(form["usanceadvance_2"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usanceadvance_2:" + ex.Message);
            }
            //貨到付款 : 總金額比例(合約)
            try
            {
                item.USANCE_GOODS_RATIO = decimal.Parse(form["usancegoods"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usancegoods:" + ex.Message);
            }
            //貨到付款 : 請款 現金比例
            try
            {
                item.USANCE_GOODS_CASH_RATIO = decimal.Parse(form["usancegoods_cash"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usancegoods_cash:" + ex.Message);
            }
            //貨到付款 : 支票 2 比例
            try
            {
                item.USANCE_GOODS_1_RATIO = decimal.Parse(form["usancegoods_1"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usancegoods_1:" + ex.Message);
            }
            //貨到付款 : 支票 3 比例
            try
            {
                item.USANCE_GOODS_2_RATIO = decimal.Parse(form["usancegoods_2"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usancegoods_2:" + ex.Message);
            }
            //完工 合約金額比例
            try
            {
                item.USANCE_FINISHED_RATIO = decimal.Parse(form["usancefinished"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usancefinished:" + ex.Message);
            }
            //完工 請款 現金比例
            try
            {
                item.USANCE_FINISHED_CASH_RATIO = decimal.Parse(form["usancefinished_cash"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usancefinished_cash:" + ex.Message);
            }
            //完工 請款 支票2比例
            try
            {
                item.USANCE_FINISHED_1_RATIO = decimal.Parse(form["usancefinished_1"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usancefinished_1:" + ex.Message);
            }
            //完工 請款 支票3比例
            try
            {
                item.USANCE_FINISHED_2_RATIO = decimal.Parse(form["usancefinished_2"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usancefinished_2:" + ex.Message);
            }
            // 尾款金額比例(合約)
            try
            {
                item.USANCE_RETENTION_RATIO = decimal.Parse(form["usanceretention"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usanceretention:" + ex.Message);
            }
            //尾款 請款現金比例
            try
            {
                item.USANCE_RETENTION_CASH_RATIO = decimal.Parse(form["usanceretention_cash"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usanceretention_cash:" + ex.Message);
            }
            //尾款 支票 2 比例
            try
            {
                item.USANCE_RETENTION_1_RATIO = decimal.Parse(form["usanceretention_1"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usanceretention_1:" + ex.Message);
            }
            //尾款 支票 3 比例
            try
            {
                item.USANCE_RETENTION_2_RATIO = decimal.Parse(form["usanceretention_2"]);
            }
            catch (Exception ex)
            {
                log.Error(item.CONTRACT_ID + " not usanceretention_2:" + ex.Message);
            }
            #endregion

            item.REMARK = form["remark"];
            SYS_USER loginUser = (SYS_USER)Session["user"];
            item.CREATE_ID = loginUser.USER_ID;
            item.CREATE_DATE = DateTime.Now;
            PurchaseFormService service = new PurchaseFormService();
            int i = service.updatePaymentTerms(item);
            if (i == 0) { msg = service.message; }
            return msg;
        }

        List<PlanItem4Map> planitems = null;
        //取得材料採購遺漏項目
        public ActionResult PendingItems(string id)
        {
            log.Info("start project id=" + id);
            PurchaseFormService service = new PurchaseFormService();
            planitems = service.getPendingItems(id);
            ViewBag.SearchResult = "共取得" + planitems.Count + "筆資料";
            return View(planitems);
        }
        public ActionResult LeadTimeMainPage(string id)
        {
            log.Info("purchase form index : projectid=" + id);
            ViewBag.projectId = id;
            SelectListItem empty = new SelectListItem();
            empty.Value = "";
            empty.Text = "";
            //取得材料合約供應商資料
            List<SelectListItem> selectMain = new List<SelectListItem>();
            foreach (string itm in service.getSupplierForContract(id))
            {
                log.Debug("supplier=" + itm);
                SelectListItem selectI = new SelectListItem();
                selectI.Value = itm;
                selectI.Text = itm;
                if (null != itm && "" != itm)
                {
                    selectMain.Add(selectI);
                }
            }
            // selectMain.Add(empty);
            ViewBag.supplier = selectMain;
            return View();
        }
        // POST : Search
        [HttpPost]
        public ActionResult LeadTimeMainPage(FormCollection f)
        {
            log.Info("projectid=" + Request["projectid"] + ",textCode1=" + Request["textCode1"] + ",textCode2=" + Request["textCode2"]);
            //加入刪除註記 預設 "N"
            List<topmeperp.Models.PlanItem4Map> lstProject = service.getPlanItem(Request["chkEx"], Request["projectid"], Request["textCode1"], Request["textCode2"], Request["textSystemMain"], Request["textSystemSub"], Request["formName"], Request["supplier"], "N", "Y");
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            ViewBag.projectId = Request["projectid"];
            SelectListItem empty = new SelectListItem();
            empty.Value = "";
            empty.Text = "";
            //取得材料合約供應商資料
            List<SelectListItem> selectMain = new List<SelectListItem>();
            foreach (string itm in service.getSupplierForContract(Request["projectid"]))
            {
                log.Debug("supplier=" + itm);
                SelectListItem selectI = new SelectListItem();
                selectI.Value = itm;
                selectI.Text = itm;
                if (null != itm && "" != itm)
                {
                    selectMain.Add(selectI);
                }
            }
            // selectMain.Add(empty);
            ViewBag.supplier = selectMain;
            return View("LeadTimeMainPage", lstProject);
        }
        //更新採購項目之採購前置時間
        public String AddLeadTime(FormCollection form)
        {
            log.Info("form:" + form.Count);
            SYS_USER u = (SYS_USER)Session["user"];
            string msg = "";
            string[] lstplanitemid = form.Get("planitemid").Split(',');
            string[] lstLeadTime = form.Get("leadtime").Split(',');
            List<PLAN_ITEM> lstItem = new List<PLAN_ITEM>();
            for (int j = 0; j < lstLeadTime.Count(); j++)
            {
                PLAN_ITEM item = new PLAN_ITEM();
                item.PROJECT_ID = form["projectId"];
                if (lstLeadTime[j].ToString() == "")
                {
                    item.LEAD_TIME = null;
                }
                else
                {
                    item.LEAD_TIME = decimal.Parse(lstLeadTime[j]);
                }
                log.Info("Lead Time =" + item.LEAD_TIME);
                item.PLAN_ITEM_ID = lstplanitemid[j];
                item.MODIFY_USER_ID = u.USER_ID;
                log.Debug("Item Project id =" + item.PROJECT_ID + "且plan item id 為" + item.PLAN_ITEM_ID + "其項目採購前置時間為" + item.LEAD_TIME);
                lstItem.Add(item);
            }
            int i = service.updateLeadTime(form["projectId"], lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新採購前置時間成功，PROJECT_ID =" + Request["projectId"];
            }

            log.Info("Request:PROJECT_ID =" + Request["projectId"]);
            return msg;
        }

        //取得廠商資料
        public string aotoCompleteData()
        {
            List<string> ls = null;
            log.Debug("get supplier");
            InquiryFormService service = new InquiryFormService();
            ls = service.getSupplier();
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            return serializer.Serialize(ls);
        }

        //取得廠商聯絡人資料
        public string getContactor()
        {
            List<TND_SUP_CONTACT_INFO> ls = null;
            log.Debug("get contact By suppliername:" + Request["Supplier"] + ", " + Request["Supplier"].Substring(0, 7));
            SupplierManage s = new SupplierManage();
            string supid = Request["Supplier"].Substring(0, 7);
            ls = s.getContactBySupplier(supid);
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            return serializer.Serialize(ls);
        }

        List<PlanItem4Map> planitems4wage = null;
        //取得工資採購遺漏項目
        public ActionResult PendingItems4Wage(string id)
        {
            log.Info("start project id=" + id);
            PurchaseFormService service = new PurchaseFormService();
            planitems4wage = service.getPendingItems4Wage(id);
            ViewBag.SearchResult = "共取得" + planitems4wage.Count + "筆資料";
            return View(planitems4wage);
        }

        List<PURCHASE_ORDER> formOutOfContract = null;
        //取得尚未發包之分項詢價單資料
        public ActionResult PendingFormTemplate(string id)
        {
            log.Info("start project id=" + id);
            //PurchaseFormService service = new PurchaseFormService();
            ContractModels contract = new ContractModels();
            formOutOfContract = service.getFormTempOutOfContractByProject(id);
            contract.planOrder = formOutOfContract;
            ViewBag.projectid = id;
            ViewBag.SearchResult = "共取得" + formOutOfContract.Count + "筆資料";
            return View(contract);
        }
        public String UpdateConStatus(FormCollection form)
        {
            log.Info("form:" + form.Count);
            string projectId = form.Get("projectid").Trim();
            log.Info("Delete PLAN_CONTRACT_PROCESS By Project ID");
            service.delAllContractByProject(projectId);
            string msg = "";
            SYS_USER u = (SYS_USER)Session["user"];
            //取得合約簽訂狀態資料
            string[] lstItemId = form.Get("contractid").Split(',');
            string[] lstBrand = form.Get("brand").Split(',');
            string[] lstProduction = form.Get("production").Split(',');
            string[] lstDate = form.Get("Date_Deliver").Split(',');
            //string[] lstRemark = form.Get("remark").Split(',');
            log.Info("select count:" + lstItemId.Count());
            var j = 0;
            for (j = 0; j < lstItemId.Count(); j++)
            {
                log.Info("item_list return No.:" + lstItemId[j]);
            }
            List<PLAN_CONTRACT_PROCESS> lstItem = new List<PLAN_CONTRACT_PROCESS>();
            for (int i = 0; i < lstItemId.Count(); i++)
            {
                PLAN_CONTRACT_PROCESS item = new PLAN_CONTRACT_PROCESS();
                item.CONTRACT_ID = lstItemId[i];
                if (lstBrand[i].ToString() == "")
                {
                    item.MATERIAL_BRAND = null;
                }
                else
                {
                    item.MATERIAL_BRAND = lstBrand[i];
                }
                if (lstProduction[i].ToString() == "")
                {
                    item.CONTRACT_PRODUCTION = null;
                }
                else
                {
                    item.CONTRACT_PRODUCTION = lstProduction[i];
                }
                if (lstDate[i].ToString() == "")
                {
                    item.DELIVERY_DATE = null;
                }
                else
                {
                    item.DELIVERY_DATE = Convert.ToDateTime(lstDate[i]);
                }
                item.CREATE_ID = u.USER_ID;
                log.Debug("Contract Id=" + item.CONTRACT_ID + ", Brand =" + item.MATERIAL_BRAND + ", Production =" + item.CONTRACT_PRODUCTION + ", deliveryDate =" + item.DELIVERY_DATE);
                lstItem.Add(item);
            }
            int k = service.AddContractProcess(projectId, lstItem);
            if (k == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新合約管制表成功";
            }
            return msg;
        }
    }
}

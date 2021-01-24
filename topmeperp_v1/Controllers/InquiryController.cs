using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;
using System.IO;
using System.Data;
using System.Web.Script.Serialization;

namespace topmeperp.Controllers
{
    public class InquiryController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(InquiryController));
        InquiryFormService service = new InquiryFormService();
        // GET: Inquiry
        public ActionResult Index(string id)
        {
            log.Info("inquiry index : projectid=" + id);
            ViewBag.projectId = id;
            SelectListItem empty = new SelectListItem();
            empty.Value = "";
            empty.Text = "";
            List<SelectListItem> selectMain = UtilService.getMainSystem(id, service);
            // selectMain.Add(empty);
            ViewBag.SystemMain = selectMain;
            List<SelectListItem> selectSub = UtilService.getSubSystem(id, service);
            //selectSub.Add(empty);
            ViewBag.SystemSub = selectSub;
            return View();
        }
        // POST : Search
        [HttpPost]
        public ActionResult Index(FormCollection f)
        {
            TnderProjectService s = new TnderProjectService();
            log.Info("projectid=" + Request["projectid"] + ",textCode1=" + Request["textCode1"] + ",textCode2=" + Request["textCode2"]);
            //加入刪除註記 預設 "N"
            List<topmeperp.Models.TND_PROJECT_ITEM> lstProject = s.getProjectItem(null, Request["projectid"], Request["textCode1"], Request["textCode2"], Request["SystemMain"], Request["systemSub"], "N");
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            ViewBag.projectId = Request["projectid"];
            //加入系統次系統
            SelectListItem empty = new SelectListItem();
            empty.Value = "";
            empty.Text = "";
            List<SelectListItem> selectMain = UtilService.getMainSystem(Request["projectid"], service);
            // selectMain.Add(empty);
            ViewBag.SystemMain = selectMain;
            List<SelectListItem> selectSub = UtilService.getSubSystem(Request["projectid"], service);
            //selectSub.Add(empty);
            ViewBag.SystemSub = selectSub;
            return View("Index", lstProject);
        }
        //Create Project Form
        [HttpPost]
        public ActionResult Create(TND_PROJECT_FORM qf)
        {
            //取得專案編號
            log.Info("Project Id:" + Request["prjId"]);
            //取得專案名稱
            log.Info("Project Name:" + Request["prjId"]);
            //取得使用者勾選品項ID
            log.Info("item_list:" + Request["chkItem"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            log.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                log.Info("item_list return No.:" + lstItemId[i]);
            }
            //建立空白詢價單
            log.Info("create new form template");
            TnderProjectService s = new TnderProjectService();
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            qf.FORM_NAME = Request["formname"];
            qf.PROJECT_ID = Request["prjId"];
            qf.CREATE_ID = u.USER_ID;
            qf.CREATE_DATE = DateTime.Now;
            qf.OWNER_NAME = uInfo.USER_NAME;
            qf.OWNER_EMAIL = uInfo.EMAIL;
            qf.OWNER_TEL = uInfo.TEL + "-" + uInfo.TEL_EXT;
            qf.OWNER_FAX = uInfo.FAX;
            TND_PROJECT_FORM_ITEM item = new TND_PROJECT_FORM_ITEM();
            string fid = s.newForm(qf, lstItemId);
            return Redirect("SinglePrjForm/" + fid);
        }
        //顯示單一詢價單、報價單功能
        public ActionResult SinglePrjForm(string id)
        {
            log.Info("http get mehtod:" + id);
            InquiryFormDetail singleForm = new InquiryFormDetail();
            service.getInqueryForm(id);
            singleForm.prjForm = service.formInquiry;
            singleForm.prjFormItem = service.formInquiryItem;
            singleForm.prj = service.getProjectById(singleForm.prjForm.PROJECT_ID);
            log.Debug("Project ID:" + singleForm.prj.PROJECT_ID);
            //取得供應商資料
            SelectListItem empty = new SelectListItem();
            empty.Value = "";
            empty.Text = "";
            return View(singleForm);
        }

        public String UpdatePrjForm(FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得供應商詢價單資料
            TND_PROJECT_FORM fm = new TND_PROJECT_FORM();
            SYS_USER loginUser = (SYS_USER)Session["user"];

            fm.PROJECT_ID = form.Get("projectid").Trim();
            //廠商資料
            if (null != form.Get("supplier") && "" != form.Get("supplier"))
            {
                fm.SUPPLIER_ID = Request["supplier"].Substring(7).Trim();
                if (form.Get("inputdateline") != "")
                {
                    fm.DUEDATE = Convert.ToDateTime(form.Get("inputdateline"));
                }
                TND_SUPPLIER s = service.getSupplierInfo(form.Get("supplier").Substring(0, 7).Trim());
                fm.CONTACT_NAME = s.CONTACT_NAME;
                fm.CONTACT_EMAIL = s.CONTACT_EMAIL;
            }
            //業務區塊
            fm.FORM_ID = form.Get("inputformnumber").Trim();
            fm.OWNER_NAME = form.Get("inputowner").Trim();
            fm.OWNER_TEL = form.Get("inputphone").Trim();
            fm.OWNER_FAX = form.Get("inputownerfax").Trim();
            fm.OWNER_EMAIL = form.Get("inputowneremail").Trim();
            fm.FORM_NAME = form.Get("formname").Trim();
            fm.ISWAGE = "N";
            if (null != form.Get("isWage"))
            {
                fm.ISWAGE = form.Get("isWage").Trim();
            }

            fm.CREATE_ID = loginUser.USER_ID;
            fm.CREATE_DATE = DateTime.Now;
            //明細區塊
            string[] lstItemId = form.Get("formitemid").Split(',');
            log.Info("select count:" + lstItemId.Count());
            var j = 0;
            for (j = 0; j < lstItemId.Count(); j++)
            {
                log.Info("item_list return No.:" + lstItemId[j]);
            }
            string fid = service.addNewSupplierForm(fm, lstItemId);
            string[] lstProjectItem = form.Get("project_item_id").Split(',');
            string[] lstPrice = form.Get("formunitprice").Split(',');
            string[] lstRemark = form.Get("remark").Split(',');
            List<TND_PROJECT_FORM_ITEM> lstItem = new List<TND_PROJECT_FORM_ITEM>();
            for (int i = 0; i < lstItemId.Count(); i++)
            {
                TND_PROJECT_FORM_ITEM item = new TND_PROJECT_FORM_ITEM();
                item.PROJECT_ITEM_ID = lstProjectItem[i];
                if (lstRemark[i].ToString() == "")
                {
                    item.ITEM_REMARK = null;
                }
                else
                {
                    item.ITEM_REMARK = lstRemark[i];
                }
                if (lstPrice[i].ToString() == "")
                {
                    item.ITEM_UNIT_PRICE = null;
                }
                else
                {
                    item.ITEM_UNIT_PRICE = decimal.Parse(lstPrice[i]);
                }
                log.Debug("Project Item Id=" + item.PROJECT_ITEM_ID + ", Price =" + item.ITEM_UNIT_PRICE);
                lstItem.Add(item);
            }
            int k = service.refreshSupplierFormItem(fid, lstItem);
            //產生廠商詢價單實體檔案
            //service.getInqueryForm(fid);
            //InquiryFormToExcel poi = new InquiryFormToExcel();
            //poi.exportExcel(service.formInquiry, service.formInquiryItem, false);
            if (fid == "")
            {
                msg = service.message;
            }
            else
            {
                msg = "新增詢價單成功";
            }

            log.Info("Request:FORM_NAME=" + form["formname"]);
            return msg;
        }
        //更新廠商詢價單資料
        public String RefreshPrjForm(string id, FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得供應商詢價單資料
            TND_PROJECT_FORM fm = new TND_PROJECT_FORM();
            SYS_USER loginUser = (SYS_USER)Session["user"];
            fm.PROJECT_ID = form.Get("projectid").Trim();
            fm.ISWAGE = "N";
            if (null != form.Get("isWage"))
            {
                fm.ISWAGE = form.Get("isWage");
            }
            fm.SUPPLIER_ID = form.Get("supplier").Trim();
            fm.DUEDATE = Convert.ToDateTime(form.Get("inputdateline"));
            fm.OWNER_NAME = form.Get("inputowner").Trim();
            fm.OWNER_TEL = form.Get("inputphone").Trim();
            fm.OWNER_FAX = form.Get("inputownerfax").Trim();
            fm.OWNER_EMAIL = form.Get("inputowneremail").Trim();
            fm.CONTACT_NAME = form.Get("inputcontact").Trim();
            fm.CONTACT_EMAIL = form.Get("inputemail").Trim();
            fm.FORM_ID = form.Get("inputformnumber").Trim();
            fm.FORM_NAME = form.Get("formname").Trim();
            fm.CREATE_ID = form.Get("createid").Trim();
            try
            {
                fm.CREATE_DATE = Convert.ToDateTime(form.Get("createdate"));
            }
            catch (Exception ex)
            {
                log.Error(ex.StackTrace);
            }
            fm.MODIFY_ID = loginUser.USER_ID;
            fm.MODIFY_DATE = DateTime.Now;
            string formid = form.Get("inputformnumber").Trim();

            string[] lstItemId = form.Get("formitemid").Split(',');
            string[] lstPrice = form.Get("formunitprice").Split(',');
            string[] lstRemark = form.Get("remark").Split(',');
            List<TND_PROJECT_FORM_ITEM> lstItem = new List<TND_PROJECT_FORM_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                TND_PROJECT_FORM_ITEM item = new TND_PROJECT_FORM_ITEM();
                item.FORM_ITEM_ID = int.Parse(lstItemId[j]);
                if (lstRemark[j].ToString() == "")
                {
                    item.ITEM_REMARK = null;
                }
                else
                {
                    item.ITEM_REMARK = lstRemark[j];
                }
                if (lstPrice[j].ToString() == "")
                {
                    item.ITEM_UNIT_PRICE = null;
                }
                else
                {
                    item.ITEM_UNIT_PRICE = decimal.Parse(lstPrice[j]);
                }
                log.Debug("Item No=" + item.FORM_ITEM_ID + "Remark =" + item.ITEM_REMARK + "Price =" + item.ITEM_UNIT_PRICE);
                lstItem.Add(item);
            }
            int i = service.refreshSupplierForm(formid, fm, lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新供應商詢價單成功，FORM_ID =" + formid;
            }

            log.Info("Request: FORM_ID = " + formid + "FORM_NAME =" + form["formname"] + "SUPPLIER_NAME =" + form["supplier"]);
            return msg;
        }

        ///詢價單管理頁面
        public ActionResult InquiryMainPage(string id)
        {
            log.Info("queryInquiry by projectID=" + id + ",status=" + Request["status"]);
            InquiryFormModel formData = new InquiryFormModel();
            if (null != id && id != "")
            {
                ViewBag.projectid = id;
                TND_PROJECT p = service.getProjectById(id);
                ViewBag.projectName = p.PROJECT_NAME;
                formData.tndTemplateProjectForm = service.getFormTemplateByProject(id);
                formData.tndProjectFormFromSupplier = service.getFormByProject(id, Request["status"]);
            }
            return View(formData);
        }
        //上傳廠商報價單
        public string FileUpload(HttpPostedFileBase file)
        {
            log.Info("Upload form from supplier:" + Request["projectid"]);
            string projectid = Request["projectid"];
            string iswage = "N";
            if (null != Request["isWage"])
            {
                log.Debug("isWage:" + Request["isWage"]);
                iswage = "Y";
            }
            int i = 0;
            //上傳至廠商報價單目錄
            if (null != file && file.ContentLength != 0)
            {
                var fileName = Path.GetFileName(file.FileName);
                string dir = ContextService.strUploadPath + "/" + projectid + "/" + ContextService.quotesFolder;
                ZipFileCreator.CreateDirectory(dir);
                var path = Path.Combine(dir, fileName);
                file.SaveAs(path);
                log.Info("Parser Excel File Begin:" + file.FileName);
                InquiryFormToExcel quoteFormService = new InquiryFormToExcel();
                try
                {
                    quoteFormService.convertInquiry2Project(path, projectid, iswage);
                }
                catch (Exception ex)
                {
                    log.Error(ex.StackTrace);
                }
                //如果詢價單編號為空白，新增詢價單資料，否則更新相關詢價單資料-new
                log.Debug("Parser Excel File Finish!");
                if (null != quoteFormService.form.FORM_ID && quoteFormService.form.FORM_ID != "")
                {
                    log.Info("Update Form for Inquiry:" + quoteFormService.form.FORM_ID);
                    i = service.refreshSupplierForm(quoteFormService.form.FORM_ID, quoteFormService.form, quoteFormService.formItems);
                }
                else
                {
                    log.Info("Create New Form for Inquiry:");
                    i = service.createInquiryFormFromSupplier(quoteFormService.form, quoteFormService.formItems);
                }
                log.Info("add supplier form record count=" + i);
            }
            if (i == -1)
            {
                return "檔案匯入失敗!!";
            }
            else
            {
                return "檔案匯入成功!!";
            }
        }
        //比價功能資料頁
        public ActionResult ComparisonMain(string id)
        {
            //傳入專案編號，
            log.Info("start project id=" + id);

            //取得專案基本資料fc
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
        //取得比價資料
        [HttpPost]
        public ActionResult ComparisonData(FormCollection form)
        {
            //傳入查詢條件
            log.Info("start project id=" + Request["id"] + ",TypeCode1=" + Request["typeCode1"] + ",typecode2=" + Request["typeCode2"] + ",SystemMain=" + Request["SystemMain"] + ",Sytem Sub=" + Request["SystemSub"] + ",Form Name=" + Request["formName"]);
            string iswage = "N";
            //取得備標品項與詢價資料
            try
            {
                if (null != Request["formName"] && "" != Request["formName"])
                {
                    if (null != Request["isWage"])
                    {
                        iswage = Request["isWage"];
                    }
                    DataTable dt = service.getComparisonDataToPivot(Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"], iswage, Request["formName"]);
                    @ViewBag.ResultMsg = "共" + dt.Rows.Count + "筆";
                    string htmlString = "<table class='table table-bordered'><tr>";
                    //處理表頭
                    for (int i = 1; i < 6; i++)
                    {
                        log.Debug("column name=" + dt.Columns[i].ColumnName);
                        htmlString = htmlString + "<th>" + dt.Columns[i].ColumnName + "</th>";
                    }
                    //處理供應商表頭
                    Dictionary<string, COMPARASION_DATA> dirSupplierQuo = service.dirSupplierQuo;
                    log.Debug("Column Count=" + dt.Columns.Count);
                    for (int i = 7; i < dt.Columns.Count; i++)
                    {
                        log.Debug("column name=" + dt.Columns[i].ColumnName);
                        string[] tmpString = dt.Columns[i].ColumnName.Split('|');
                        //<a href="/Inquiry/SinglePrjForm/@item.FORM_ID" target="_blank">@item.FORM_ID</a>
                        decimal tAmount = (decimal)dirSupplierQuo[tmpString[1]].TAmount;
                        string strAmout = string.Format("{0:C0}", tAmount);

                        htmlString = htmlString + "<th><table><tr><td>" + tmpString[0] + '(' + tmpString[2] + ')' +
                           "<br/><button type='button' class='btn-xs' onclick=\"clickSupplier('" + tmpString[1] + "','" + iswage + "')\"><span class='glyphicon glyphicon-ok' aria-hidden='true'></span></button>" +
                            "<button type='button' class='btn-xs'><a href='/Inquiry/SinglePrjForm/" + tmpString[1] + "'" + " target='_blank'><span class='glyphicon glyphicon-list-alt' aria-hidden='true'></span></a></button>" +
                            "<button type='button' class='btn-xs' onclick=\"chaneFormStatus('" + tmpString[1] + "','註銷')\"><span class='glyphicon glyphicon-remove' aria-hidden='true'></span></a></button>" +
                            "</td><tr><td style='text-align:center;background-color:yellow;' >" + strAmout + "</td></tr></table></th>";
                    }
                    htmlString = htmlString + "</tr>";
                    //處理資料表
                    foreach (DataRow dr in dt.Rows)
                    {
                        htmlString = htmlString + "<tr>";
                        for (int i = 1; i < 5; i++)
                        {
                            htmlString = htmlString + "<td>" + dr[i] + "</td>";
                        }
                        //單價欄位  <input type='text' id='cost_@item.PROJECT_ITEM_ID' name='cost_@item.PROJECT_ITEM_ID' size='5' />
                        //decimal price = decimal.Parse(dr[5].ToString());
                        if (dr[5].ToString() != "")
                        {
                            log.Debug("data row col 5=" + (decimal)dr[5]);
                            htmlString = htmlString + "<td><input type='text' id='cost_" + dr[1] + "' name='cost_" + dr[1] + "' size='5' value='" + String.Format("{0:N0}", (decimal)dr[5]) + "' /></td>";
                        }
                        else
                        {
                            htmlString = htmlString + "<td></td>";
                        }
                        //String.Format("{0:C}", 0);
                        //處理報價資料
                        for (int i = 7; i < dt.Columns.Count; i++)
                        {
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
                    DataTable dt = service.getComparisonDataToPivot(Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"], iswage, Request["formName"]);
                    @ViewBag.ResultMsg = "共" + dt.Rows.Count + "筆";
                    string htmlString = "<table class='table table-bordered'><tr>";
                    //處理表頭
                    for (int i = 1; i < 6; i++)
                    {
                        log.Debug("column name=" + dt.Columns[i].ColumnName);
                        htmlString = htmlString + "<th>" + dt.Columns[i].ColumnName + "</th>";
                    }
                    //處理供應商表頭
                    Dictionary<string, COMPARASION_DATA> dirSupplierQuo = service.dirSupplierQuo;
                    log.Debug("Column Count=" + dt.Columns.Count);
                    for (int i = 6; i < dt.Columns.Count; i++)
                    {
                        log.Debug("column name=" + dt.Columns[i].ColumnName);
                        string[] tmpString = dt.Columns[i].ColumnName.Split('|');
                        //<a href="/Inquiry/SinglePrjForm/@item.FORM_ID" target="_blank">@item.FORM_ID</a>
                        decimal tAmount = (decimal)dirSupplierQuo[tmpString[1]].TAmount;
                        string strAmout = string.Format("{0:C0}", tAmount);

                        htmlString = htmlString + "<th><table><tr><td>" + tmpString[0] + '(' + tmpString[2] + ')' +
                            "<br/><button type='button' class='btn-xs' onclick=\"clickSupplier('" + tmpString[1] + "','" + iswage + "')\"><span class='glyphicon glyphicon-ok' aria-hidden='true'></span></button>" +
                            "<button type='button' class='btn-xs'><a href='/Inquiry/SinglePrjForm/" + tmpString[1] + "'" + " target='_blank'><span class='glyphicon glyphicon-list-alt' aria-hidden='true'></span></a></button>" +
                            "<button type='button' class='btn-xs' onclick=\"chaneFormStatus('" + tmpString[1] + "','註銷')\"><span class='glyphicon glyphicon-remove' aria-hidden='true'></span></a></button>" +
                            "</td><tr><td style='text-align:center;background-color:yellow;' >" + strAmout + "</td></tr></table></th>";
                    }
                    htmlString = htmlString + "</tr>";
                    //處理資料表
                    foreach (DataRow dr in dt.Rows)
                    {
                        htmlString = htmlString + "<tr>";
                        for (int i = 1; i < 5; i++)
                        {
                            htmlString = htmlString + "<td>" + dr[i] + "</td>";
                        }
                        //單價欄位  <input type='text' id='cost_@item.PROJECT_ITEM_ID' name='cost_@item.PROJECT_ITEM_ID' size='5' />
                        //decimal price = decimal.Parse(dr[5].ToString());
                        if (dr[5].ToString() != "")
                        {
                            log.Debug("data row col 5=" + (decimal)dr[5]);
                            htmlString = htmlString + "<td><input type='text' id='cost_" + dr[1] + "' name='cost_" + dr[1] + "' size='5' value='" + String.Format("{0:N0}", (decimal)dr[5]) + "' /></td>";
                        }
                        else
                        {
                            htmlString = htmlString + "<td></td>";
                        }
                        //String.Format("{0:C}", 0);
                        //處理報價資料
                        for (int i = 6; i < dt.Columns.Count; i++)
                        {
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
                    //產生畫面
                    IHtmlString str = new HtmlString(htmlString);
                    ViewBag.htmlString = str;
                }
            }
            catch (Exception e)
            {
                log.Error(e.StackTrace);
                ViewBag.htmlString = e.Message;
            }
            return PartialView();
        }
        //更新單項成本資料
        public string UpdateCost4Item()
        {
            log.Info("ProjectItemID=" + Request["pitmid"] + ",Cost=" + Request["price"] + ",iswage=" + Request["iswage"]);

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
        //依據詢價單內容，更新標單所有單價
        public string BatchUpdateCost(string formid)
        {
            log.Info("formid=" + Request["formid"] + ",iswage=" + Request["iswage"]);
            int i = service.batchUpdateCostFromQuote(Request["formid"], Request["iswage"]);
            return "更新成功!!";
        }
        //成本分析
        public ActionResult costAnalysis(string id)
        {
            //產生成本分析Excel 並以固定檔案供使用者下載使用
            ViewBag.projectid = id;
            log.Info("Cost Analysis for projectid=" + id);
            CostAnalysisOutput excel = new CostAnalysisOutput();
            excel.exportExcel(id);
            ViewBag.url = "/UploadFile/" + id + "/" + id + "_CostAnalysis.xlsx";
            ViewBag.ErrorMsg = excel.errorMessage;
            return View();
        }
        //批次產生空白詢價單
        public string createEmptyForm()
        {
            log.Info("project id=" + Request["projectid"]);
            SYS_USER u = (SYS_USER)Session["user"];
            int i = service.createEmptyForm(Request["projectid"], u);
            return "共產生 " + i + "空白詢價單樣本!!";
        }
        /// <summary>
        /// 下載空白詢價單
        /// </summary>
        public void downLoadInquiryForm()
        {
            string formid = Request["formid"];
            service.getInqueryForm(formid);
            if (null != service.formInquiry)
            {
                InquiryFormToExcel poi = new InquiryFormToExcel();
                //檔案位置
                string fileLocation = poi.exportExcel(service.formInquiry, service.formInquiryItem, false);
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
        public void changeStatus()
        {
            string formid = Request["formId"];
            string status = Request["status"];
            log.Debug("change form status:" + formid + ",status=" + status);
            service.changeProjectFormStatus(formid, status);
        }
        public ActionResult mailForm()
        {
            log.Debug("test email sender!");
            return View();
        }
        public string sendMain()
        {
            ///取得寄件者資料
            string strSenderAddress = Request["textSenderAddress"];
            ///取得收件者資料，有逗號結尾，所以將其移除
            string strReceiveAddress = Request["textReceiveAddress"];
            if (strReceiveAddress.EndsWith(","))
            {
                strReceiveAddress = strReceiveAddress.Substring(0, strReceiveAddress.Length - 1);
            }
            log.Info("send email:strSenderAddress=" + strSenderAddress + ",strReceiveAddress=" + strReceiveAddress);
            // List<string> MailList = new List<string>();
            //MailList.Add(strReceiveAddress);
            //取得主旨資料
            string strSubject = Request["textSubject"];
            //取得內容
            string strContent = Request["textContent"];
            log.Debug("test email sender!");
            EMailService es = new EMailService();
            //設定附件檔案
            string projectid = Request["textProjectId"];
            string poid = Request["textPOID"];
            string realFilePath = ContextService.strUploadPath + "\\" + projectid + "\\" + ContextService.quotesFolder + "\\" + poid + ".xlsx";
            SYS_USER u = (SYS_USER)Session["user"];
            log.Debug("Attachment file path=" + realFilePath);
            if (es.SendMailByGmail(strSenderAddress, u.USER_NAME, strReceiveAddress, null, strSubject, strContent, realFilePath))
            {
                return "發送成功!!";
            }
            else
            {
                return "發送失敗!!";
            }
        }
        //取得廠商資料
        public string aotoCompleteData()
        {
            List<string> ls = null;
            log.Debug("get supplier");
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
    }
}

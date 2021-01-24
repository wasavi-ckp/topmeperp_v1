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
using System.Reflection;

namespace topmeperp.Controllers
{
    public class MaterialManageController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(InquiryController));
        PurchaseFormService service = new PurchaseFormService();
        ProjectPlanService planService = new ProjectPlanService();

        // GET: MaterialManage
        [topmeperp.Filter.AuthFilter]
        public ActionResult Index()
        {
            List<topmeperp.Models.TND_PROJECT> lstProject = SearchProjectByName("", "專案執行");
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            return View(lstProject);
        }

        private List<topmeperp.Models.TND_PROJECT> SearchProjectByName(string projectname, string status)
        {
            if (projectname != null)
            {
                log.Info("search project by 名稱 =" + projectname);
                List<topmeperp.Models.TND_PROJECT> lstProject = new List<TND_PROJECT>();
                using (var context = new topmepEntities())
                {
                    lstProject = context.TND_PROJECT.SqlQuery("select * from TND_PROJECT p "
                        + "where p.PROJECT_NAME Like '%' + @projectname + '%' AND STATUS=@status;",
                         new SqlParameter("projectname", projectname), new SqlParameter("status", status)).ToList();
                }
                log.Info("get project count=" + lstProject.Count);
                return lstProject;
            }
            else
            {
                return null;
            }
        }

        //Plan Task 任務連結物料
        public ActionResult PlanTask(string id)
        {
            log.Debug("show sreen for apply for material");
            ViewBag.projectid = id;
            TnderProjectService tndservice = new TnderProjectService();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.TreeString = planService.getProjectTask4Tree(id);

            //SelectListItem empty = new SelectListItem();
            //empty.Value = "";
            //empty.Text = "";
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
            return View();
        }
        public string getTypeCodeL2()
        {
            string typecode1 = Request["typecode1"];
            log.Debug("get type code1=" + typecode1);
            TypeManageService typeService = new TypeManageService();
            List<REF_TYPE_MAIN> lstType1 = typeService.getTypeMainL2(typecode1);
            string strOpt = "{";

            for (int idx = 0; idx < lstType1.Count; idx++)
            {
                REF_TYPE_MAIN itm = lstType1[idx];
                log.Debug("REF_TYPE_MAIN=" + idx + "," + lstType1[idx].TYPE_DESC);
                if (idx == lstType1.Count - 1)
                {
                    strOpt = strOpt + "\"" + itm.TYPE_CODE_1 + itm.TYPE_CODE_2 + "\":\"" + itm.TYPE_DESC + "\"}";
                }
                else
                {
                    strOpt = strOpt + "\"" + itm.TYPE_CODE_1 + itm.TYPE_CODE_2 + "\":\"" + itm.TYPE_DESC + "\",";
                }
            }

            // System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            // string itemJson = objSerializer.Serialize("");
            log.Debug("REF_TYPE_MAIN=" + strOpt);
            return strOpt;
        }
        public string getSubType()
        {
            string typecode = Request["typecode"];
            log.Debug("get type code=" + typecode);
            TypeManageService typeService = new TypeManageService();
            List<REF_TYPE_SUB> lstType = typeService.getSubType(typecode);
            string strOpt = "{\"0\":\"全部\",";

            for (int idx = 0; idx < lstType.Count; idx++)
            {
                REF_TYPE_SUB itm = lstType[idx];
                log.Debug("REF_TYPE_SUB=" + idx + "," + lstType[idx].TYPE_DESC);
                if (idx == lstType.Count - 1)
                {
                    strOpt = strOpt + "\"" + itm.SUB_TYPE_CODE + "\":\"" + itm.TYPE_DESC + "\"}";
                }
                else
                {
                    strOpt = strOpt + "\"" + itm.SUB_TYPE_CODE + "\":\"" + itm.TYPE_DESC + "\",";
                }
            }

            // System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            // string itemJson = objSerializer.Serialize("");
            log.Debug("REF_TYPE_SUB=" + strOpt);
            return strOpt;
        }
        //物料申購
        public ActionResult Application(FormCollection form)
        {
            log.Info("Access to Application page!!");
            ViewBag.projectid = form["projectid"];
            TnderProjectService tndservice = new TnderProjectService();
            TND_PROJECT p = tndservice.getProjectById(form["projectid"]);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.applyDate = DateTime.Now;
            string[] deviceItemId = null;
            string[] fpItemId = null;
            string[] pepItemId = null;
            string[] lcpItemId = null;
            string[] pluItemId = null;
            string[] fwItemId = null;
            string[] notMapItemId = null;
            List<string> AllItemId = new List<string>();
            //取得使用者勾選任務ID
            //設備資料
            if (null != form["map_device"])
            {
                log.Info("device task_list:" + Request["map_device"]);
                deviceItemId = Request["map_device"].ToString().Split(',');

                log.Info("select count:" + deviceItemId.Count());
                var i = 0;
                for (i = 0; i < deviceItemId.Count(); i++)
                {
                    log.Info("device task_list return No.:" + deviceItemId[i]);
                    AllItemId.Add(deviceItemId[i]);
                    //ViewBag.uid = lstItemId[i];
                }
            }
            //消防電資料
            if (null != form["map_fp"])
            {
                log.Info("fp task_list:" + Request["map_fp"]);
                fpItemId = Request["map_fp"].ToString().Split(',');

                log.Info("select count:" + fpItemId.Count());
                var i = 0;
                for (i = 0; i < fpItemId.Count(); i++)
                {
                    log.Info("fp task_list return No.:" + fpItemId[i]);
                    AllItemId.Add(fpItemId[i]);
                }
            }
            //電氣資料
            if (null != form["map_pep"])
            {
                log.Info("pep task_list:" + Request["map_pep"]);
                pepItemId = Request["map_pep"].ToString().Split(',');

                log.Info("select count:" + pepItemId.Count());
                var i = 0;
                for (i = 0; i < pepItemId.Count(); i++)
                {
                    log.Info("pep task_list return No.:" + pepItemId[i]);
                    AllItemId.Add(pepItemId[i]);
                }
            }
            //弱電資料
            if (null != form["map_lcp"])
            {
                log.Info("pep task_list:" + Request["map_lcp"]);
                lcpItemId = Request["map_lcp"].ToString().Split(',');

                log.Info("select count:" + lcpItemId.Count());
                var i = 0;
                for (i = 0; i < lcpItemId.Count(); i++)
                {
                    log.Info("lcp task_list return No.:" + lcpItemId[i]);
                    AllItemId.Add(lcpItemId[i]);
                }
            }
            //給排水
            if (null != form["map_plu"])
            {
                log.Info("plu task_list:" + Request["map_plu"]);
                pluItemId = Request["map_plu"].ToString().Split(',');

                log.Info("select count:" + pluItemId.Count());
                var i = 0;
                for (i = 0; i < pluItemId.Count(); i++)
                {
                    log.Info("plu task_list return No.:" + pluItemId[i]);
                    AllItemId.Add(pluItemId[i]);
                }
            }
            //消防水
            if (null != form["map_fw"])
            {
                log.Info("fw task_list:" + Request["map_fw"]);
                fwItemId = Request["map_fw"].ToString().Split(',');

                log.Info("select count:" + fwItemId.Count());
                var i = 0;
                for (i = 0; i < fwItemId.Count(); i++)
                {
                    log.Info("fw task_list return No.:" + fwItemId[i]);
                    AllItemId.Add(fwItemId[i]);
                }
            }
            //非圖算數量
            if (null != form["item_not_map"])
            {
                log.Info("device not in map list:" + Request["item_not_map"]);
                notMapItemId = Request["item_not_map"].ToString().Split(',');

                log.Info("select count:" + notMapItemId.Count());
                var i = 0;
                for (i = 0; i < notMapItemId.Count(); i++)
                {
                    log.Info("device task_list return No.:" + notMapItemId[i]);
                    AllItemId.Add(notMapItemId[i]);
                    //ViewBag.uid = lstItemId[i];
                }
            }

            if (null == form["map_device"] && null == form["map_fp"] && null == form["map_fw"] 
                && null == form["map_pep"] && null == form["map_lcp"] && null == form["map_plu"]
                && null == form["item_not_map"])
            {
                TempData["result"] = "沒有選取要申購的項目名稱，請重新查詢後並勾選物料項目!";
                return Redirect("PlanTask?id=" + form["projectid"]);
            }
            else
            {
                List<PurchaseRequisition> lstPR = service.getPurchaseItemByMap(form["projectid"], AllItemId);
                //補收件人住址等資料
                SYS_USER u = (SYS_USER)Session["user"];
                ViewBag.recipient = u.USER_NAME;
                ViewBag.location = p.LOCATION;
                return View(lstPR);
            }
        }

        [HttpPost]
        //儲存申購單(申購單草稿)
        public ActionResult SavePR(PLAN_PURCHASE_REQUISITION pr)
        {
            //取得專案編號
            log.Info("Project Id:" + Request["id"]);
            //取得專案名稱
            log.Info("Project Name:" + Request["projectName"]);
            //取得使用者勾選品項ID
            log.Info("item_list:" + Request["chkItem"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            log.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                log.Info("item_list return No.:" + lstItemId[i]);
            }
            string[] Qty = Request["need_qty"].Split(',');
            string[] Date = Request["Date_${index}"].Split(',');
            string[] Remark = Request["remark"].Split(',');
            //建立申購單
            log.Info("create new Purchase Requisition");
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            pr.PROJECT_ID = Request["id"];
            pr.CREATE_USER_ID = u.USER_ID;
            pr.CREATE_DATE = DateTime.Now;
            pr.RECIPIENT = Request["recipient"];
            pr.LOCATION = Request["location"];
            pr.REMARK = Request["caution"];
            pr.MEMO = Request["memo"];
            pr.STATUS = 0; //表示申購單未送出，只是存檔而已
            //pr.PRJ_UID = int.Parse(Request["prj_uid"]);
            PLAN_PURCHASE_REQUISITION_ITEM item = new PLAN_PURCHASE_REQUISITION_ITEM();
            string prid = service.newPR(Request["id"], pr, lstItemId);
            List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_PURCHASE_REQUISITION_ITEM items = new PLAN_PURCHASE_REQUISITION_ITEM();
                items.PLAN_ITEM_ID = lstItemId[j];
                if (Qty[j].Trim() != "")
                {
                    items.NEED_QTY = decimal.Parse(Qty[j]);
                }
                try
                {
                    items.NEED_DATE = Convert.ToDateTime(Date[j]);
                }
                catch (Exception ex)
                {
                    log.Error(ex.StackTrace);
                }
                items.REMARK = Remark[j];
                log.Debug("Item No=" + items.PLAN_ITEM_ID + ", Qty =" + items.NEED_QTY + ", Date =" + items.NEED_DATE);
                lstItem.Add(items);
            }
            int k = service.refreshPR(prid, pr, lstItem);
            return Redirect("SinglePR?id=" + prid + "&prjid=" + pr.PROJECT_ID);
        }
        //申購單查詢
        public ActionResult PurchaseRequisition(string id)
        {
            log.Info("Search For Purchase Requisition !!");
            ViewBag.projectid = id;
            TnderProjectService tndservice = new TnderProjectService();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            List<PRFunction> lstPR = service.getPRByPrjId(id, Request["create_date"], Request["keyname"], Request["prid"], Request["status"]);
            return View(lstPR);
        }

        public ActionResult Search()
        {
            //log.Info("projectid=" + Request["id"] + ", keyname =" + Request["keyname"] + ", prid =" + Request["prid"] + ", create_id =" + Request["create_date"] + ", status =" + int.Parse(Request["status"]));
            string status = Request["status"];
            List<PRFunction> lstPR = service.getPRByPrjId(Request["id"], Request["create_date"], Request["keyname"], Request["prid"], status);
            ViewBag.SearchResult = "共取得" + lstPR.Count + "筆資料";
            ViewBag.projectId = Request["id"];
            ViewBag.projectName = Request["projectName"];
            return View("PurchaseRequisition", lstPR);
        }

        //顯示單一申購單功能
        public ActionResult SinglePR(string id, string prjid)
        {
            log.Info("http get mehtod:" + id);
            PurchaseRequisitionDetail singleForm = new PurchaseRequisitionDetail();
            string parentId = "";
            if (null == prjid)
            {
                prjid = "";
            }
            service.getPRByPrId(id, parentId, prjid);
            singleForm.planPR = service.formPR;
            singleForm.planPRItem = service.PRItem;
            singleForm.prj = service.getProjectById(singleForm.planPR.PROJECT_ID);
            log.Debug("Project ID:" + singleForm.prj.PROJECT_ID);
            return View(singleForm);
        }

        //新增申購單物料品項
        public String addPRFormItem(FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "更新成功!!";

            PLAN_PURCHASE_REQUISITION_ITEM item = new PLAN_PURCHASE_REQUISITION_ITEM();
            item.PR_ID = form["dia_form_id"];
            item.ITEM_DESC = form["dia_item_desc"];
            item.ITEM_UNIT = form["dia_item_unit"];
            try
            {
                item.NEED_QTY = decimal.Parse(form["dia_need_quantity"]);
            }
            catch (Exception ex)
            {
                log.Error(item.ITEM_DESC + " not need qty:" + ex.Message);
            }
            try
            {
                item.NEED_DATE = Convert.ToDateTime(form["dia_need_date"]);
            }
            catch (Exception ex)
            {
                log.Error(item.NEED_DATE + " not need date:" + ex.Message);
            }
            item.REMARK = form["dia_item_remark"];
            int i = service.addPRItem(item);
            return msg + "(" + i + ")";
        }

        //更新申購單草稿
        public String RefreshPR(string id, FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得申購單資料
            PLAN_PURCHASE_REQUISITION pr = new PLAN_PURCHASE_REQUISITION();
            SYS_USER u = (SYS_USER)Session["user"];
            pr.PROJECT_ID = form.Get("projectid").Trim();
            //pr.PRJ_UID = int.Parse(form.Get("prjuid").Trim());
            pr.STATUS = 0;
            pr.PR_ID = form.Get("pr_id").Trim();
            pr.RECIPIENT = form.Get("recipient").Trim();
            pr.LOCATION = form.Get("location").Trim();
            pr.REMARK = form.Get("caution").Trim();
            pr.MEMO = form.Get("memo").Trim();
            try
            {
                pr.MESSAGE = form.Get("message").Trim();
            }
            catch (Exception ex)
            {
                log.Error(ex.StackTrace);
            }
            pr.CREATE_USER_ID = u.USER_ID;
            pr.MODIFY_DATE = DateTime.Now;
            try
            {
                pr.CREATE_DATE = Convert.ToDateTime(form.Get("apply_date"));
            }
            catch (Exception ex)
            {
                log.Error(ex.StackTrace);
            }
            string formid = form.Get("pr_id").Trim();
            string[] lstItemId = form.Get("pr_item_id").Split(',');
            string[] lstQty = form.Get("need_qty").Split(',');
            string[] lstRemark = form.Get("remark").Split(',');
            string[] lstDate = form.Get("date").Split(',');
            List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_PURCHASE_REQUISITION_ITEM item = new PLAN_PURCHASE_REQUISITION_ITEM();
                item.PR_ITEM_ID = int.Parse(lstItemId[j]);
                if (lstQty[j].ToString() == "")
                {
                    item.NEED_QTY = null;
                }
                else
                {
                    item.NEED_QTY = decimal.Parse(lstQty[j]);
                }
                log.Debug("Item No=" + item.PR_ITEM_ID + ", Need Qty =" + item.NEED_QTY);
                item.REMARK = lstRemark[j];
                if (lstDate[j] != "")
                {
                    item.NEED_DATE = DateTime.ParseExact(lstDate[j], "yyyy/MM/dd", CultureInfo.InvariantCulture);
                }
                lstItem.Add(item);
            }
            int i = service.updatePR(formid, pr, lstItem,null);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新申購單草稿成功，PR_ID =" + formid;
            }

            log.Info("Request: PR_ID = " + formid + "Task Id =" + form["prjuid"]);
            return msg;
        }

        //新增申購單
        public String CreatePR(string id, FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得申購單資料
            PLAN_PURCHASE_REQUISITION pr = new PLAN_PURCHASE_REQUISITION();
            SYS_USER loginUser = (SYS_USER)Session["user"];
            pr.PROJECT_ID = form.Get("projectid").Trim();
            //pr.PRJ_UID = int.Parse(form.Get("prjuid").Trim());
            pr.STATUS = 10;
            pr.PR_ID = form.Get("pr_id").Trim();
            pr.RECIPIENT = form.Get("recipient").Trim();
            pr.LOCATION = form.Get("location").Trim();
            pr.REMARK = form.Get("caution").Trim();
            pr.MEMO = form.Get("memo").Trim();
            pr.CREATE_USER_ID = loginUser.USER_ID;
            pr.MODIFY_DATE = DateTime.Now;
            try
            {
                pr.CREATE_DATE = Convert.ToDateTime(form.Get("apply_date")); 
            }
            catch (Exception ex)
            {
                log.Error(ex.StackTrace);
            }
            string formid = form.Get("pr_id").Trim();
            string[] lstItemId = form.Get("pr_item_id").Split(',');
            string[] lstQty = form.Get("need_qty").Split(',');
            string[] lstRemark = form.Get("remark").Split(',');
            string[] lstDate = form.Get("date").Split(',');
            List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_PURCHASE_REQUISITION_ITEM item = new PLAN_PURCHASE_REQUISITION_ITEM();
                item.PR_ITEM_ID = int.Parse(lstItemId[j]);
                if (lstQty[j].ToString() == "")
                {
                    item.NEED_QTY = null;
                }
                else
                {
                    item.NEED_QTY = decimal.Parse(lstQty[j]);
                }
                log.Debug("Item No=" + item.PR_ITEM_ID + ", Need Qty =" + item.NEED_QTY);
                item.REMARK = lstRemark[j];
                if (lstDate[j] != "")
                {
                    item.NEED_DATE = DateTime.ParseExact(lstDate[j], "yyyy/MM/dd", CultureInfo.InvariantCulture);
                }
                lstItem.Add(item);
            }
            //建立申購單須送出email
            int i = service.updatePR(formid, pr, lstItem, loginUser);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "新增申購單成功，PR_ID =" + formid;
            }

            log.Info("Request: PR_ID = " + formid + "Task Id =" + form["prjuid"]);
            return msg;
        }

        //更新申購單狀態(退件:代碼5,不處理)
        public string changePRStatus(FormCollection form)
        {
            log.Info("formid=" + form["pr_id"]);
            int i = service.changePRStatus(form["pr_id"], form["message"]);
            return "更新成功!!";
        }
        //採購作業
        public ActionResult PurchaseOrder(string id)
        {
            log.Info("Access to Purchase Order Page !!");
            ViewBag.projectid = id;
            TnderProjectService tndservice = new TnderProjectService();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            List<PurchaseOrderFunction> lstPO = service.getPRBySupplier(id);
            return View(lstPO);
        }
        //取得申購單之供應商合約項目
        public ActionResult PurchaseOperation(string prjid, string prid, string sup)
        {
            log.Info("Access to Purchase Operation page!!");
            ViewBag.projectid = prjid;
            TnderProjectService tndservice = new TnderProjectService();
            TND_PROJECT p = tndservice.getProjectById(prjid);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.supplier = sup;
            ViewBag.parentPrId = prid;
            ViewBag.OrderDate = DateTime.Now;
            PurchaseRequisitionDetail singleForm = new PurchaseRequisitionDetail();
            service.getPRByPrId(ViewBag.parentPrId, ViewBag.parentPrId, prjid);
            singleForm.planPR = service.formPR;
            ViewBag.recipient = singleForm.planPR.RECIPIENT;
            ViewBag.location = singleForm.planPR.LOCATION;
            ViewBag.caution = singleForm.planPR.REMARK;
            //ViewBag.status = singleForm.planPR.STATUS;
            List<PurchaseRequisition> lstPR = service.getPurchaseItemBySupplier(String.Join("-", ViewBag.parentPrId, ViewBag.supplier), prjid);
            return View(lstPR);
        }
        //新增採購單
        public ActionResult AddPO(PLAN_PURCHASE_REQUISITION pr)
        {
            //取得專案編號
            log.Info("Project Id:" + Request["id"]);
            //取得專案名稱
            log.Info("Project Name:" + Request["projectName"]);
            //取得使用者勾選品項ID
            log.Info("item_list:" + Request["chkItem"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            log.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                log.Info("item_list return No.:" + lstItemId[i]);
            }
            string[] Qty = Request["order_qty"].Split(',');
            //string[] PlanItemId = Request["planitemid"].Split(',');
            List<string> lstQty = new List<string>();
            //List<string> lstPlanItemId = new List<string>();
            var m = 0;
            for (m = 0; m < Qty.Count(); m++)
            {
                if (Qty[m] != "" && null != Qty[m])
                {
                    lstQty.Add(Qty[m]);
                }
            }
            //建立採購單
            log.Info("create new Purchase Order");
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            //SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            pr.PROJECT_ID = Request["id"];
            pr.CREATE_USER_ID = u.USER_ID;
            pr.CREATE_DATE = DateTime.Now;
            pr.RECIPIENT = Request["recipient"];
            pr.LOCATION = Request["location"];
            pr.REMARK = Request["caution"];
            pr.SUPPLIER_ID = Request["supplier"];
            pr.PARENT_PR_ID = Request["parent_pr_id"];
            pr.STATUS = 20;
            pr.MEMO = Request["memo"];
            log.Debug("memo = " + Request["memo"]);
            pr.MESSAGE = Request["message"];
            log.Debug("message = " + Request["message"]);
            PLAN_PURCHASE_REQUISITION_ITEM item = new PLAN_PURCHASE_REQUISITION_ITEM();
            string prid = service.newPO(Request["id"], pr, lstItemId, Request["parent_pr_id"]);
            List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_PURCHASE_REQUISITION_ITEM items = new PLAN_PURCHASE_REQUISITION_ITEM();
                items.PLAN_ITEM_ID = lstItemId[j];
                items.ORDER_QTY = decimal.Parse(lstQty[j]);
                log.Debug("Item No=" + items.PLAN_ITEM_ID + ", Qty =" + items.ORDER_QTY);
                lstItem.Add(items);
            }
            int k = service.refreshPO(prid, pr, lstItem,u);
            return Redirect("SinglePO?id=" + prid + "&prjid=" + Request["id"]);
        }

        //採購單查詢
        public ActionResult PurchaseOrderIndex(string id)
        {
            log.Info("Search For Purchase Order !!");
            ViewBag.projectid = id;
            TnderProjectService tndservice = new TnderProjectService();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            return View();
        }

        [HttpPost]
        public ActionResult PurchaseOrderIndex(FormCollection f)
        {
            log.Info("projectid=" + Request["id"] + ", supplier =" + Request["supplier"] + ", prid =" + Request["prid"] + ", create_id =" + Request["create_date"] + ", parent_prid =" + Request["parent_prid"] + ", keyname =" + Request["keyname"]);
            List<PRFunction> lstPO = service.getPOByPrjId(Request["id"], Request["create_date"], Request["supplier"], Request["prid"], Request["parent_prid"], Request["keyname"]);
            ViewBag.SearchResult = "共取得" + lstPO.Count + "筆資料";
            ViewBag.projectId = Request["id"];
            ViewBag.projectName = Request["projectName"];
            return View("PurchaseOrderIndex", lstPO);
        }

        //顯示單一採購單功能
        public ActionResult SinglePO(string id, string prjid)
        {
            log.Info("http get prid :" + id + ",and project id = " + prjid);
            PurchaseRequisitionDetail singleForm = new PurchaseRequisitionDetail();
            string parentId = "";
            service.getPRByPrId(id, parentId, prjid);
            singleForm.planPR = service.formPR;
            singleForm.planPRItem = service.PRItem;
            ViewBag.status = singleForm.planPR.STATUS;
            log.Debug("planPR:" + singleForm.planPR.PR_ID +",PARENT_ID="+ singleForm.planPR.PARENT_PR_ID + ",status=" + singleForm.planPR.STATUS);
            ViewBag.orderDate = singleForm.planPR.CREATE_DATE.Value.ToString("yyyy/MM/dd");
            singleForm.prj = service.getProjectById(singleForm.planPR.PROJECT_ID);
            log.Debug("Project ID:" + singleForm.prj.PROJECT_ID);

            ViewBag.prId = singleForm.planPR.PARENT_PR_ID;//service.getParentPrIdByPrId(id);
            return View(singleForm);
        }

        //更新採購單資料
        public String RefreshPO(string id, FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得採購單資料
            PLAN_PURCHASE_REQUISITION pr = new PLAN_PURCHASE_REQUISITION();
            SYS_USER loginUser = (SYS_USER)Session["user"];
            pr.PROJECT_ID = form.Get("projectid").Trim();
            pr.PR_ID = form.Get("pr_id").Trim();
            pr.RECIPIENT = form.Get("recipient").Trim();
            pr.LOCATION = form.Get("location").Trim();
            pr.REMARK = form.Get("caution").Trim();
            pr.SUPPLIER_ID = form.Get("supplier").Trim();
            pr.PARENT_PR_ID = form.Get("parent_pr_id").Trim();
            pr.STATUS = int.Parse(form.Get("status").Trim());
            pr.CREATE_USER_ID = loginUser.USER_ID;
            pr.MODIFY_DATE = DateTime.Now;
            pr.MEMO = form.Get("memo").Trim();
            pr.MESSAGE = form.Get("message").Trim();
            try
            {
                pr.CREATE_DATE = Convert.ToDateTime(form.Get("order_date"));
            }
            catch (Exception ex)
            {
                log.Error(ex.StackTrace);
            }
            string formid = form.Get("pr_id").Trim();
            string[] lstItemId = form.Get("pr_item_id").Split(',');
            string[] lstQty = form.Get("order_qty").Split(',');

            List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_PURCHASE_REQUISITION_ITEM item = new PLAN_PURCHASE_REQUISITION_ITEM();
                item.PR_ITEM_ID = int.Parse(lstItemId[j]);
                if (lstQty[j].ToString() == "")
                {
                    item.ORDER_QTY = null;
                }
                else
                {
                    item.ORDER_QTY = decimal.Parse(lstQty[j]);
                }
                log.Debug("Item No=" + item.PR_ITEM_ID + ", Order Qty =" + item.ORDER_QTY);
                lstItem.Add(item);
            }
            int i = service.updatePO(formid, pr, lstItem);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新採購單成功，PR_ID =" + formid;
            }

            log.Info("Request: PR_ID = " + formid + " 供應商名稱=" + form["supplier"]);
            return msg;
        }

        //更新採購單備忘錄
        public string RefreshMemo(FormCollection form)
        {
            log.Info("formid=" + form["pr_id"]);
            int i = service.changeMemo(form["pr_id"], form["memo"]);
            return "更新成功!!";
        }
        //驗收作業
        public ActionResult Receipt(string id, string prjid)
        {
            log.Info("http get mehtod:" + id);
            PurchaseRequisitionDetail singleForm = getReceiveForm(id, prjid);
            return View(singleForm);
        }
        //列印驗收單
        public ActionResult ReceiptPrint(string id, string prjid)
        {
            log.Info("http get mehtod:" + id);
            PurchaseRequisitionDetail singleForm = getReceiveForm(id, prjid);
            return View(singleForm);
        }
        private PurchaseRequisitionDetail getReceiveForm(string id, string prjid)
        {
            //取得驗收單資料
            PurchaseRequisitionDetail singleForm = new PurchaseRequisitionDetail();
            service.getPRByPrId(id, id, prjid);
            singleForm.planPR = service.formPR;
            singleForm.planPRItem = service.PRItem;
            ViewBag.receiptDate = DateTime.Now.ToString("yyyy/MM/dd");
            singleForm.prj = service.getProjectById(singleForm.planPR.PROJECT_ID);
            log.Debug("Project ID:" + singleForm.prj.PROJECT_ID);
            ViewBag.EnablePrint = "";
            return singleForm;
        }

        //新增驗收單資料
        public ActionResult AddReceipt(PLAN_PURCHASE_REQUISITION pr)
        {
            //取得專案編號
            log.Info("Project Id:" + Request["projectid"]);
            //取得專案名稱
            log.Info("Project Name:" + Request["projectName"]);
            //取得使用者勾選品項ID
            log.Info("item_list:" + Request["chkItem"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            log.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                log.Info("item_list return No.:" + lstItemId[i]);
            }
            string[] Qty = Request["receipt_qty"].Split(',');
            //string[] PlanItemId = Request["planitemid"].Split(',');
            List<string> lstQty = new List<string>();
            //List<string> lstPlanItemId = new List<string>();
            var m = 0;
            for (m = 0; m < Qty.Count(); m++)
            {
                if (Qty[m] != "" && null != Qty[m])
                {
                    lstQty.Add(Qty[m]);
                }
            }
            //建立驗收單
            log.Info("create new Receipt");
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            pr.PROJECT_ID = Request["projectid"];
            pr.CREATE_USER_ID = u.USER_ID;
            pr.CREATE_DATE = DateTime.Now;
            pr.RECIPIENT = Request["recipient"];
            pr.LOCATION = Request["location"];
            pr.REMARK = Request["caution"];
            pr.SUPPLIER_ID = Request["supplier"];
            pr.PARENT_PR_ID = Request["pr_id"];
            
            pr.STATUS = 30;
            pr.MEMO = Request["memo"];
            pr.MESSAGE = Request["message"];
            PLAN_PURCHASE_REQUISITION_ITEM item = new PLAN_PURCHASE_REQUISITION_ITEM();
            string prid = service.newRP(Request["projectid"], pr, lstItemId, Request["pr_id"]);
            //string allKey = prid + '-' + Request["pr_id"];
            List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_PURCHASE_REQUISITION_ITEM items = new PLAN_PURCHASE_REQUISITION_ITEM();
                items.PLAN_ITEM_ID = lstItemId[j];
                items.RECEIPT_QTY = decimal.Parse(lstQty[j]);
                log.Debug("Item No=" + items.PLAN_ITEM_ID + ", Qty =" + items.RECEIPT_QTY);
                lstItem.Add(items);
            }
            string closeFlag= Request["closeFlag"]; 
            int k = service.refreshRP(prid, pr, lstItem, closeFlag);
            return Redirect("SingleRP?id=" + prid + "&parentId=" + Request["pr_id"] + "&prjid=" + Request["projectid"]);
        }

        //顯示單一驗收單功能
        public ActionResult SingleRP(string id, string parentId, string prjid)
        {
            log.Info("http get mehtod:" + id);
            PurchaseRequisitionDetail singleForm = new PurchaseRequisitionDetail();
            service.getPRByPrId(id, parentId, prjid);
            singleForm.planPR = service.formPR;
            singleForm.planPRItem = service.PRItem;
            ViewBag.receiptDate = singleForm.planPR.CREATE_DATE.Value.ToString("yyyy/MM/dd");
            singleForm.prj = service.getProjectById(singleForm.planPR.PROJECT_ID);
            log.Debug("Project ID:" + singleForm.prj.PROJECT_ID);
            ViewBag.EnablePrint = "True";
            return View("Receipt", singleForm);
        }
        //驗收單明細
        public ActionResult ReceiptList(string id)
        {
            log.Info("Access to Receipt List !!");
            List<PRFunction> lstRP = service.getRPByPrId(id);
            return View(lstRP);
        }

        //庫存查詢
        public ActionResult InventoryIndex(string id)
        {
            log.Info("Search For Inventory of All Item !!");
            ViewBag.projectid = id;
            TnderProjectService tndservice = new TnderProjectService();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            SYS_USER u = (SYS_USER)Session["user"];
            ViewBag.createid = u.USER_ID;
            //SelectListItem empty = new SelectListItem();
            //empty.Value = "";
            //empty.Text = "";
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

            TypeManageService typeService = new TypeManageService();
            List<REF_TYPE_MAIN> lstType1 = typeService.getTypeMainL1();

            //取得九宮格
            List<SelectListItem> selectType1 = new List<SelectListItem>();
            for (int idx = 0; idx < lstType1.Count; idx++)
            {
                log.Debug("REF_TYPE_MAIN=" + idx + "," + lstType1[idx].CODE_1_DESC);
                SelectListItem selectI = new SelectListItem();
                selectI.Value = lstType1[idx].TYPE_CODE_1;
                selectI.Text = lstType1[idx].CODE_1_DESC;
                selectType1.Add(selectI);
            }
            ViewBag.TypeCodeL1 = selectType1;
            List<PurchaseRequisition> lstItem = service.getInventoryByPrjId(id, Request["item"], Request["TypeCodeL2"], Request["TypeSub"], Request["systemMain"], Request["systemSub"], Request["remark"]);
            return View(lstItem);
        }

        public ActionResult SearchInventory()
        {
            //log.Info("projectid=" + Request["id"] + ", planitemname =" + Request["item"] + ", TypeCodeL1 =" + Request["TypeCodeL1"] + ", TypeCodeL2 =" + Request["TypeCodeL2"] + ", TypeSub =" + Request["TypeSub"] + ", systemMain =" + Request["systemMain"] + ", systemSub =" + Request["systemSub"]);
            List<PurchaseRequisition> lstItem = service.getInventoryByPrjId(Request["id"], Request["item"], Request["TypeCodeL2"], Request["TypeSub"], Request["systemMain"], Request["systemSub"], Request["remark"]);
            ViewBag.SearchResult = "共取得" + lstItem.Count + "筆資料";
            ViewBag.projectId = Request["id"];
            ViewBag.projectName = Request["projectName"];
            return View("InventoryIndex", lstItem);
        }

        //領料作業
        public ActionResult DeliveryIndex(string id)
        {
            log.Info("Access To Delivery Operation Page，Project Id =" + id);
            ViewBag.projectid = id;
            TnderProjectService tndservice = new TnderProjectService();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            SYS_USER u = (SYS_USER)Session["user"];
            ViewBag.createid = u.USER_ID;
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
            ViewBag.SystemSub = selectSub;
            TypeManageService typeService = new TypeManageService();
            List<REF_TYPE_MAIN> lstType1 = typeService.getTypeMainL1();
            //取得九宮格
            List<SelectListItem> selectType1 = new List<SelectListItem>();
            for (int idx = 0; idx < lstType1.Count; idx++)
            {
                log.Debug("REF_TYPE_MAIN=" + idx + "," + lstType1[idx].CODE_1_DESC);
                SelectListItem selectI = new SelectListItem();
                selectI.Value = lstType1[idx].TYPE_CODE_1;
                selectI.Text = lstType1[idx].CODE_1_DESC;
                selectType1.Add(selectI);
            }
            ViewBag.TypeCodeL1 = selectType1;
            List<PurchaseRequisition> lstItem = service.getInventoryByPrjId(id, Request["item"], Request["TypeCodeL2"], Request["TypeSub"], Request["systemMain"], Request["systemSub"], Request["remark"]);
            return View(lstItem);
        }

        public ActionResult SearchDelivery()
        {
            //log.Info("projectid=" + Request["id"] + ", planitemname =" + Request["item"] + ", TypeCodeL1 =" + Request["TypeCodeL1"] + ", TypeCodeL2 =" + Request["TypeCodeL2"] + ", TypeSub =" + Request["TypeSub"] + ", systemMain =" + Request["systemMain"] + ", systemSub =" + Request["systemSub"]);
            List<PurchaseRequisition> lstItem = service.getInventoryByPrjId(Request["id"], Request["item"], Request["TypeCodeL2"], Request["TypeSub"], Request["systemMain"], Request["systemSub"], Request["remark"]);
            ViewBag.SearchResult = "共取得" + lstItem.Count + "筆資料";
            ViewBag.projectId = Request["id"];
            ViewBag.projectName = Request["projectName"];
            return View("DeliveryIndex", lstItem);
        }
        //新增物料提領資料
        public ActionResult AddDelivery(FormCollection form)
        {
            //取得專案編號
            log.Info("Project Id:" + Request["prjId"]);
            //取得專案名稱
            log.Info("Project Name:" + Request["prjName"]);
            //取得使用者勾選品項ID
            log.Info("item_list:" + Request["chkItem"]);
            string[] lstItemId = Request["chkItem"].ToString().Split(',');
            log.Info("select count:" + lstItemId.Count());
            var i = 0;
            for (i = 0; i < lstItemId.Count(); i++)
            {
                log.Info("item_list return No.:" + lstItemId[i]);
            }
            string[] Qty = Request["delivery_qty"].Split(',');
            //string[] lstPlanItemId = Request["planitemid"].Split(',');
            List<string> lstQty = new List<string>();
            var m = 0;
            for (m = 0; m < Qty.Count(); m++)
            {
                if (Qty[m] != "" && null != Qty[m])
                {
                    lstQty.Add(Qty[m]);
                }
            }
            PLAN_PURCHASE_REQUISITION pr = new PLAN_PURCHASE_REQUISITION();
            SYS_USER loginUser = (SYS_USER)Session["user"];
            pr.PROJECT_ID = form.Get("prjId").Trim();
            //pr.PRJ_UID = int.Parse(form.Get("prjuid").Trim());
            pr.STATUS = 40;
            pr.CREATE_USER_ID = loginUser.USER_ID;
            pr.CREATE_DATE = DateTime.Now;
            pr.CAUTION = form.Get("caution").Trim();
            pr.RECIPIENT = form.Get("recipient").Trim();
            PLAN_ITEM_DELIVERY item = new PLAN_ITEM_DELIVERY();
            string deliveryorderid = service.newDelivery(Request["prjId"], pr, lstItemId, loginUser.USER_ID);
            List<PLAN_ITEM_DELIVERY> lstItem = new List<PLAN_ITEM_DELIVERY>();
            for (int j = 0; j < lstItemId.Count(); j++)
            {
                PLAN_ITEM_DELIVERY items = new PLAN_ITEM_DELIVERY();
                items.PLAN_ITEM_ID = lstItemId[j].Split('/')[0];
                items.REMARK = lstItemId[j].Split('/')[1];
                items.DELIVERY_QTY = decimal.Parse(lstQty[j]);
                log.Debug("Item No=" + items.PLAN_ITEM_ID + ", Remark =" + items.REMARK + ", Qty =" + items.DELIVERY_QTY);
                lstItem.Add(items);
            }
            int k = service.refreshDelivery(deliveryorderid, lstItem);
            return Redirect("SingleDO?id=" + deliveryorderid + "&prjid=" + Request["prjId"]);
        }

        //物料進出明細
        public ActionResult DeliveryItem(string id)
        {
            log.Info("Access to Delivery Item !!");
            List<PurchaseRequisition> lstItem = service.getDeliveryByItemId(id);
            log.Debug("plan_item_id = " + id);
            return View(lstItem);
        }
        //驗收單查詢
        public ActionResult ReceiptSearch(string id)
        {
            log.Info("Search For Receipt !!");
            ViewBag.projectid = id;
            TnderProjectService tndservice = new TnderProjectService();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            return View();
        }

        [HttpPost]
        public ActionResult ReceiptSearch(FormCollection f)
        {
            log.Info("projectid=" + Request["id"] + ", keyword =" + Request["keyword"]);
            List<PRFunction> lstRP = service.getRP4Delivery(Request["id"], Request["keyword"]);
            ViewBag.SearchResult = "共取得" + lstRP.Count + "筆資料";
            ViewBag.projectId = Request["id"];
            ViewBag.projectName = Request["projectName"];
            return View("ReceiptSearch", lstRP);
        }

        //顯示單一領料單功能
        public ActionResult SingleDO(string id, string prjid)
        {
            log.Info("http get mehtod:" + id);
            PurchaseRequisitionDetail singleForm = new PurchaseRequisitionDetail();
            string parentId = "";
            service.getPRByPrId(id, parentId, prjid);
            singleForm.planPR = service.formPR;
            singleForm.planPRItem = service.PRItem;
            singleForm.planDOItem = service.DOItem;
            singleForm.prj = service.getProjectById(singleForm.planPR.PROJECT_ID);
            log.Debug("Project ID:" + singleForm.prj.PROJECT_ID);
            //ViewBag.prId = service.getParentPrIdByPrId(id);
            return View(singleForm);
        }
        //領料單查詢
        public ActionResult DeliverySearch(string id)
        {
            log.Info("Search For Delivery !!");
            ViewBag.projectid = id;
            TnderProjectService tndservice = new TnderProjectService();
            TND_PROJECT p = tndservice.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            return View();
        }

        [HttpPost]
        public ActionResult DeliverySearch(FormCollection f)
        {
            log.Info("projectid=" + Request["id"] + ", recipient =" + Request["recipient"] + ", prid =" + Request["prid"] + ", caution =" + Request["caution"]);
            List<PRFunction> lstDO = service.getDOByPrjId(Request["id"], Request["recipient"], Request["prid"], Request["caution"]);
            ViewBag.SearchResult = "共取得" + lstDO.Count + "筆資料";
            ViewBag.projectId = Request["id"];
            ViewBag.projectName = Request["projectName"];
            return View("DeliverySearch", lstDO);
        }

        //3天內須驗收的物料品項
        public ActionResult CheckForReceipt(string prid, string prjid, int status)
        {
            log.Info("Access to CheckForReceip Page By PR Id =" + prid);
            ViewBag.projectid = prjid;
            TnderProjectService tndservice = new TnderProjectService();
            TND_PROJECT p = tndservice.getProjectById(prjid);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.formid = prid;
            List<PurchaseRequisition> lstItem = service.getPlanItemByNeedDate(prid, prjid, status);
            ViewBag.SearchResult = "共取得" + lstItem.Count + "筆資料";
            return View(lstItem);
        }
        public String RejectPRById(FormCollection form)
        {
            //取得申購單編號
            log.Info("PR form Id:" + form["pr_id"]);
            //更新費用單狀態
            log.Info("Reject PR Form ");
            string formid = form.Get("pr_id").Trim();
            //申購單(已退件) STATUS = -10
            string msg = "";
            int i = service.RejectPRByPRId(formid);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "申購單已退回";
            }
            return msg;
        }
        //申購單驗收未完成的物料品項
        public ActionResult CheckForReceiptQty(string prid, string prjid, int status)
        {
            log.Info("Access to CheckForReceipQty Page By PR Id =" + prid);
            ViewBag.projectid = prjid;
            TnderProjectService tndservice = new TnderProjectService();
            TND_PROJECT p = tndservice.getProjectById(prjid);
            ViewBag.projectName = p.PROJECT_NAME;
            ViewBag.formid = prid;
            List<PurchaseRequisition> lstItem = service.getItemNeedReceivedByPrId(prid, prjid, status);
            ViewBag.SearchResult = "共取得" + lstItem.Count + "筆資料";
            return View(lstItem);
        }
        /// <summary>
        /// 物料採購單
        /// </summary>
        public void downLoadMaterialForm()
        {
            string key = Request["key"];
            string[] keys = key.Split('-');
            string projectid = keys[0];
            string formid = keys[1];
            string parentId = keys[2]; 
            bool isOrder = false;
            bool isDO = false;
            PlanService pservice = new PlanService();
            pservice.getProject(projectid);
            service.getPRByPrId(formid, parentId, projectid);
            if (formid.Substring(0, 1) != "D")
            {
                isOrder = true;
            }
            if (formid.Substring(0, 2) == "DO")
            {
                isDO = true;
            }
            if (null != pservice.project)
            {
                MaterialFormToExcel poi = new MaterialFormToExcel();
                //檔案位置
                string fileLocation = poi.exportExcel(pservice.project, service.formPR, service.PRItem, service.DOItem, isOrder, isDO);
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
        //刪除無用的申購單
        public string delPR()
        {
            SYS_USER u = (SYS_USER)Session["user"];
            string prid = Request["pr_id"];
            service.delPR(prid);
            log.Info(u.USER_ID + ",remove pr_id =" + prid);
            return "/MaterialManage/PurchaseRequisition/" + Request["projectid"];
        }
    }
}

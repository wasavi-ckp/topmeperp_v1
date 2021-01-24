using log4net;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;

namespace topmeperp.Controllers
{
    /// <summary>
    /// 專案任務相關功能
    /// </summary>
    public class ProjectPlanController : Controller
    {
        static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        ProjectPlanService planService = new ProjectPlanService();
        // GET: ProjectPaln
        public ActionResult Index()
        {
            if (null != Request["projectid"])
            {
                string projectid = Request["projectid"];
                string prjuid = null;
                PLAN_TASK task = null;
                log.Debug("get project task by project:" + projectid + ",roottag=" + Request["roottag"]);
                if (null != projectid && "" != projectid)
                {
                    prjuid = Request["prjuid"];
                    log.Debug("get project task by child task by prj_uid:" + prjuid);
                }

                if (null != Request["roottag"] && "Y" == Request["roottag"])
                {
                    task = planService.getRootTask(projectid);
                    if (null != task)
                    {
                        log.Debug("task=" + task.PRJ_UID);
                        prjuid = task.PRJ_UID.ToString();
                    }
                }

                DataTable dt = null;
                if (null == prjuid || prjuid == "")
                {
                    //取得所有任務
                    dt = planService.getProjectTask(projectid);
                }
                else
                {
                    //取得所有子項任務
                    dt = planService.getChildTask(projectid, int.Parse(prjuid));
                }
                string htmlString = "<table class='table table-bordered'>";

                htmlString = htmlString + "<tr><th>層級</th><th>任務名稱</th><th>開始時間</th><th>完成時間</th><th>工期</th><th>--</th><th>--</th></tr>";
                foreach (DataRow dr in dt.Rows)
                {
                    DateTime stardate = DateTime.Parse(dr[4].ToString());
                    DateTime finishdate = DateTime.Parse(dr[5].ToString());

                    htmlString = htmlString + "<tr><td>" + dr[1] + "<input type='checkbox' name='roottag' id='roottag' onclick='setRootTask(" + dr[2] + ")' /></td><td>" + dr[0] + "</td>"
                        + "<td>" + stardate.ToString("yyyy-MM-dd") + "</td><td>" + finishdate.ToString("yyyy-MM-dd") + "</td><td>" + dr[6] + "</td>"
                        + "<td ><a href =\"Index?projectid=" + projectid + "&prjuid=" + dr[3] + "\">上一層 </a></td>"
                        + "<td><a href=\"Index?projectid=" + projectid + "&prjuid=" + dr[2] + "\">下一層 </a></td></tr>";
                }
                htmlString = htmlString + "</table>";
                ViewBag.htmlResult = htmlString;
                ViewBag.projectId = Request["projectid"];
            }
            return View();
        }
        //上傳project 檔案，建立專案任務
        public ActionResult uploadFile(HttpPostedFileBase file)
        {
            //設置施工管理資料夾
            if (null != file)
            {
                log.Info("upload file!!" + file.FileName);
                string projectId = Request["projectid"];
                string projectFolder = ContextService.strUploadPath + "/" + projectId + "/" + ContextService.projectMgrFolder;
                if (Directory.Exists(projectFolder))
                {
                    //資料夾存在
                    log.Info("Directory Exist:" + projectFolder);
                }
                else
                {
                    //if directory not exist create it
                    Directory.CreateDirectory(projectFolder);
                }
                if (null != file && file.ContentLength != 0)
                {
                    //2.upload project file
                    //2.2 將上傳檔案存檔
                    var fileName = Path.GetFileName(file.FileName);
                    var path = Path.Combine(projectFolder, fileName);
                    file.SaveAs(path);
                    OfficeProjectService s = new OfficeProjectService();
                    s.convertProject(projectId, path);
                    s.import2Table();
                }
            }
            return Redirect("Index?projectid=" + Request["projectid"] + "&roottag=" + Request["roottag"]);
            // return View("Index/projectid=" + Request["projectid"]);
        }
        //設定合約範圍起始任務
        public string setRootFlag()
        {
            log.Debug("projectid=" + Request["projectid"] + ",prjuid=" + Request["prjuid"]);
            int i = planService.setRootTask(Request["projectid"], Request["prjuid"]);
            return "設定完成!!(" + i + ")";
        }
        //專案任務與圖算數量設定畫面
        public ActionResult ManageTaskDetail()
        {
            log.Debug("show sreen for task manage");
            string projectid = Request["projectid"];
            ViewBag.projectId = projectid;
            ViewBag.TreeString = planService.getProjectTask4Tree(projectid);
            Dictionary<string, object> sec = TypeSelectComponet.getMapItemQueryCriteria(projectid);
            ViewBag.SystemMain = sec["SystemMain"];
            ViewBag.SystemSub = sec["SystemSub"];
            ViewBag.TypeCodeL1 = sec["TypeCodeL1"];

            return View();
        }
        //查詢圖算資訊
        public ActionResult getMapItem4Task(FormCollection f)
        {
            string projectid, typeCode1, typeCode2, systemMain, systemSub, primeside, primesideName, secondside, secondsideName, mapno, buildno, devicename, mapType, strart_id, end_id;
            TypeSelectComponet.getMapItem(f, out projectid, out typeCode1, out typeCode2, out systemMain, out systemSub, out primeside, out primesideName, out secondside, out secondsideName, out mapno, out buildno, out devicename, out mapType, out strart_id, out end_id);
            if (null == f["mapType"] || "" == f["mapType"])
            {
                ViewBag.Message = "至少需選擇一項施作項目!!";
                return PartialView("_getMapItem4Task", null);
            }
            string[] mapTypes = mapType.Split(',');
            for (int i = 0; i < mapTypes.Length; i++)
            {
                switch (mapTypes[i])
                {
                    case "MAP_DEVICE"://設備
                        log.Debug("MapType: MAP_DEVICE(設備)");
                        //增加九宮格、次九宮格、主系統、次系統等條件
                        planService.getMapItem(projectid, devicename, strart_id, end_id, typeCode1, typeCode2, systemMain, systemSub);
                        break;
                    case "MAP_PEP"://電氣管線
                        log.Debug("MapType: MAP_PEP(電氣管線)");
                        //增加一次側名稱、二次側名稱
                        planService.getMapPEP(projectid, mapno, buildno, primeside, primesideName, secondside, secondsideName, devicename);
                        break;
                    case "MAP_LCP"://弱電管線
                        log.Debug("MapType: MAP_LCP(弱電管線)");
                        planService.getMapLCP(projectid, mapno, buildno, primeside, primesideName, secondside, secondsideName, devicename);
                        break;
                    case "TND_MAP_PLU"://給排水
                        log.Debug("MapType: TND_MAP_PLU(給排水)");
                        planService.getMapPLU(projectid, mapno, buildno, primeside, primesideName, secondside, secondsideName, devicename);
                        break;
                    case "MAP_FP"://消防電
                        log.Debug("MapType: MAP_FP(消防電)");
                        planService.getMapFP(projectid, mapno, buildno, primeside, primesideName, secondside, secondsideName, devicename);
                        break;
                    case "MAP_FW"://消防水
                        planService.getMapFW(projectid, mapno, buildno, primeside, primesideName, secondside, secondsideName, devicename);
                        log.Debug("MapType: MAP_FW(消防水)");
                        break;
                    case "NOT_MAP"://不在圖算內
                        planService.getItemNotMap(projectid, mapno, buildno, primeside, primesideName, secondside, secondsideName, devicename);
                        log.Debug("MapType: MAP_FW(消防水)");
                        break;
                    default:
                        log.Debug("MapType nothing!!");
                        break;
                }
            }
            ViewBag.Message = planService.resultMessage;
            return PartialView("_getMapItem4Task", planService.viewModel);
        }
        //上傳任務與圖算資料
        public string uploadTaskAndItem(HttpPostedFileBase file1)
        {
            string projectid = Request["id"];
            log.Debug("ProjectID=" + projectid + ",Upload ProjectItem=" + file1.FileName);
            TnderProjectService service = new TnderProjectService();
            service.getProjectById(projectid);
            SYS_USER u = (SYS_USER)Session["user"];
            if (null != file1 && file1.ContentLength != 0)
            {
                try
                {
                    //2.解析Excel
                    log.Info("Parser Excel data:" + file1.FileName);
                    //2.1 將上傳檔案存檔
                    var fileName = Path.GetFileName(file1.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    log.Info("save excel file:" + path);
                    file1.SaveAs(path);
                    //2.2 解析Excel 檔案
                    //poiservice.ConvertDataForTenderProject(prj.PROJECT_ID, (int)prj.START_ROW_NO);
                    ProjectTask2MapService poiservice = new ProjectTask2MapService();
                    poiservice.InitializeWorkbook(path);
                    poiservice.transAllSheet(projectid);
                    //List<PLAN_TASK2MAPITEM> lstTask2Map = poiservice.ConvertDataForMapFW(projectid);
                    int i = planService.createTask2Map(poiservice.lstTask2Map);
                    service.strMessage = poiservice.errorMessage;
                    //log.Debug("add PLAN_TASK2MAPITEM =" + i);
                    // service.refreshProjectItem(poiservice.lstProjectItem);
                }
                catch (Exception ex)
                {
                    log.Error(ex.StackTrace);
                    return ex.Message;
                }
            }
            if (service.strMessage != null)
            {
                return service.strMessage;
            }
            else
            {
                return "匯入成功!!";
            }
        }

        //設定任務圖算
        public string choiceMapItem(FormCollection f)
        {
            if (null == f["checkNodeId"] || "" == f["checkNodeId"])
            {
                return "請選擇專案任務!!";
            }
            if (null != f["map_device"])
            {
                log.Debug("projectId=" + f["projectid"] + ",prjuid=" + f["checkNodeId"] + ",mapids=" + f["map_device"]);
                //設備
                int i = planService.choiceMapItem(f["projectid"], f["checkNodeId"], f["map_device"]);
                log.Debug("modify records count=" + i);
            }
            if (null != f["map_pep"])
            {
                log.Debug("projectId=" + f["projectid"] + ",prjuid=" + f["checkNodeId"] + ",map_pep=" + f["map_pep"]);
                //電氣管線
                int i = planService.choiceMapItemPEP(f["projectid"], f["checkNodeId"], f["map_pep"]);
                log.Debug("modify records count=" + i);
            }
            if (null != f["map_lcp"])
            {
                log.Debug("projectId=" + f["projectid"] + ",prjuid=" + f["checkNodeId"] + ",map_lcp=" + f["map_lcp"]);
                //弱電
                int i = planService.choiceMapItemLCP(f["projectid"], f["checkNodeId"], f["map_lcp"]);
                log.Debug("modify records count=" + i);
            }

            if (null != f["map_plu"])
            {
                log.Debug("projectId=" + f["projectid"] + ",prjuid=" + f["checkNodeId"] + ",map_plu=" + f["map_plu"]);
                //給排水
                int i = planService.choiceMapItemPLU(f["projectid"], f["checkNodeId"], f["map_plu"]);
                log.Debug("modify records count=" + i);
            }

            if (null != f["map_fp"])
            {
                log.Debug("projectId=" + f["projectid"] + ",prjuid=" + f["checkNodeId"] + ",map_fp=" + f["map_fp"]);
                //消防電
                int i = planService.choiceMapItemFP(f["projectid"], f["checkNodeId"], f["map_fp"]);
                log.Debug("modify records count=" + i);
            }
            if (null != f["map_fw"])
            {
                log.Debug("projectId=" + f["projectid"] + ",prjuid=" + f["checkNodeId"] + ",map_fw=" + f["map_fw"]);
                //消防水
                int i = planService.choiceMapItemFW(f["projectid"], f["checkNodeId"], f["map_fw"]);
                log.Debug("modify records count=" + i);
            }

            return "設定成功";
        }
        public ActionResult getActionItem4Task(FormCollection f)
        {
            log.Debug("projectId=" + f["projectid"] + ",prjuid=" + f["checkNodeId"]);
            //planService
            return PartialView("_getProjecttem4Task", planService.getItemInTask(f["projectid"], f["checkNodeId"]));
        }
        //填寫日報step1 :選取任務
        public ActionResult dailyReport(string id)
        {
            if (null == id || "" == id)
            {
                id = Request["projectid"];
            }
            string strRptDate = "";
            if (null != Request["reportDate"])
            {
                strRptDate = Request["reportDate"].Trim();
            }
            DateTime dtTaskDate = DateTime.Now;
            if (strRptDate != "")
            {
                dtTaskDate = DateTime.Parse(strRptDate);
            }
            log.Debug("get Task for plan by day=" + dtTaskDate);
            List<PLAN_TASK> lstTask = planService.getTaskByDate(id, dtTaskDate);
            ViewBag.projectName = planService.getProject(id).PROJECT_NAME;
            ViewBag.projectId = id;
            ViewBag.reportDate = dtTaskDate.ToString("yyyy/MM/dd");
            return View(lstTask);
        }
        //填寫日報step2 :選取填寫內容
        public ActionResult dailyReportItem()
        {
            DailyReport dailyRpt = null;
            if (null == Request["rptID"])
            {
                ViewBag.projectId = Request["projectid"];
                ViewBag.projectName = planService.getProject(Request["projectid"]).PROJECT_NAME;
                ViewBag.prj_uid = Request["prjuid"];
                ViewBag.taskName = planService.getProjectTask(Request["projectid"], int.Parse(Request["prjuid"])).TASK_NAME;
                ViewBag.RptDate = Request["rptDate"];
                dailyRpt = planService.newDailyReport(Request["projectid"], int.Parse(Request["prjuid"]));
                ViewBag.selWeather = getDropdownList4Weather("");
            }
            else
            {
                string strRptId = Request["rptID"];
                ViewBag.RptId = strRptId;
                dailyRpt = planService.getDailyReport(strRptId);

                ViewBag.projectId = dailyRpt.dailyRpt.PROJECT_ID;

                ViewBag.projectName = planService.getProject(dailyRpt.dailyRpt.PROJECT_ID).PROJECT_NAME;
                ViewBag.prj_uid = dailyRpt.lstRptTask[0].PRJ_UID;
                ViewBag.taskName = planService.getProjectTask(dailyRpt.dailyRpt.PROJECT_ID, int.Parse(dailyRpt.lstRptTask[0].PRJ_UID.ToString())).TASK_NAME;
                ViewBag.RptDate = string.Format("{0:yyyy/MM/dd}", dailyRpt.dailyRpt.REPORT_DATE);
                ViewBag.ddlWeather = getDropdownList4Weather(dailyRpt.dailyRpt.WEATHER);
            }

            //1.依據任務取得相關施作項目內容
            return View(dailyRpt);
        }
        private List<SelectListItem> getDropdownList4Weather(string selecValue)
        {
            string[] aryWeather = { "晴", "陰", "雨" };
            List<SelectListItem> lstWeather = new List<SelectListItem>();
            for (int i = 0; i < aryWeather.Length; i++)
            {
                bool selected = aryWeather[i].Equals(selecValue);
                lstWeather.Add(new SelectListItem()
                {
                    Text = aryWeather[i],
                    Value = aryWeather[i],
                    Selected = selected
                });
            }
            return lstWeather;
        }

        //儲存日報數量紀錄
        public void saveItemRow(FormCollection f)
        {
            SYS_USER u = (SYS_USER)Session["user"];
            log.Debug("projectId=" + Request["txtProjectId"] + ",prjUid=" + Request["txtPrjUid"] + ",ReportId=" + Request["reportID"]);
            log.Debug("form Data ItemId=" + Request["planItemId"]);
            log.Debug("form Data Qty=planItemQty" + Request["planItemQty"]);

            string projectid = Request["txtProjectId"];
            int prjuid = int.Parse(Request["txtPrjUid"]);
            string strWeather = Request["selWeather"];
            string strSummary = Request["txtSummary"];
            string strSenceUser = Request["txtSenceUser"];
            string strSupervision = Request["txtSupervision"];
            string strOwner = Request["txtOwner"];
            string strRptDate = Request["reportDate"];

            DailyReport newDailyRpt = new DailyReport();

            PLAN_DALIY_REPORT RptHeader = new PLAN_DALIY_REPORT();
            RptHeader.PROJECT_ID = projectid;
            RptHeader.WEATHER = strWeather;
            RptHeader.SUMMARY = strSummary;
            RptHeader.SCENE_USER_NAME = strSenceUser;
            RptHeader.SUPERVISION_NAME = strSupervision;
            RptHeader.OWNER_NAME = strOwner;
            newDailyRpt.dailyRpt = RptHeader;
            RptHeader.REPORT_DATE = DateTime.Parse(strRptDate);
            //取得日報編號
            SerialKeyService snService = new SerialKeyService();
            if (null == Request["reportID"] || "" == Request["reportID"])
            {
                RptHeader.REPORT_ID = snService.getSerialKey(planService.KEY_ID);
                RptHeader.CREATE_DATE = DateTime.Now;
                RptHeader.CREATE_USER_ID = u.USER_ID;
            }
            else
            {
                RptHeader.REPORT_ID = Request["ReportID"];
                RptHeader.CREATE_DATE = DateTime.Parse(Request["txtCreateDate"]);
                RptHeader.CREATE_USER_ID = Request["txtCreateUserId"];
                RptHeader.MODIFY_DATE = DateTime.Now;
                RptHeader.MODIFY_USER_ID = u.USER_ID;
            }

            //建立專案任務資料 (結構是支援多項任務，僅先使用一筆)
            newDailyRpt.lstRptTask = new List<PLAN_DR_TASK>();
            PLAN_DR_TASK RptTask = new PLAN_DR_TASK();
            RptTask.PROJECT_ID = projectid;
            RptTask.PRJ_UID = prjuid;
            RptTask.REPORT_ID = RptHeader.REPORT_ID;
            newDailyRpt.lstRptTask.Add(RptTask);
            //處理料件
            newDailyRpt.lstRptItem = new List<PLAN_DR_ITEM>();
            if (null != Request["planItemId"])
            {
                string[] aryPlanItem = Request["planItemId"].Split(',');
                string[] aryPlanItemQty = Request["planItemQty"].Split(',');
                string[] aryAccumulateQty = Request["accumulateQty"].Split(',');

                log.Debug("count ItemiD=" + aryPlanItem.Length + ",qty=" + aryPlanItemQty.Length);
                newDailyRpt.lstRptItem = new List<PLAN_DR_ITEM>();
                for (int i = 0; i < aryPlanItem.Length; i++)
                {
                    PLAN_DR_ITEM item = new PLAN_DR_ITEM();
                    item.PLAN_ITEM_ID = aryPlanItem[i];
                    item.PROJECT_ID = projectid;
                    item.REPORT_ID = RptHeader.REPORT_ID;
                    if ("" != aryPlanItemQty[i])
                    {
                        item.FINISH_QTY = decimal.Parse(aryPlanItemQty[i]);
                    }
                    if ("" != aryAccumulateQty[i])
                    {
                        item.LAST_QTY = decimal.Parse(aryAccumulateQty[i]);
                    }
                    newDailyRpt.lstRptItem.Add(item);
                }
            }
            //處理出工資料
            newDailyRpt.lstWokerType4Show = new List<PLAN_DR_WORKER>();
            if (null != Request["txtSupplierId"])
            {
                ///出工廠商
                string[] arySupplier = Request["txtSupplierId"].Split(',');
                ///出工人數
                string[] aryWorkerQty = Request["txtWorkerQty"].Split(',');
                ///備註
                string[] aryRemark = Request["txtRemark"].Split(',');
                for (int i = 0; i < arySupplier.Length; i++)
                {
                    PLAN_DR_WORKER item = new PLAN_DR_WORKER();
                    item.REPORT_ID = RptHeader.REPORT_ID;
                    item.SUPPLIER_ID = arySupplier[i];
                    log.Debug("Supplier Info=" + item.SUPPLIER_ID);
                    if ("" != aryWorkerQty[i].Trim())
                    {
                        item.WORKER_QTY = decimal.Parse(aryWorkerQty[i]);
                        newDailyRpt.lstWokerType4Show.Add(item);
                    }
                    item.REMARK = aryRemark[i];
                }
                log.Debug("count WorkerD=" + arySupplier.Length);
            }
            //處理點工資料
            newDailyRpt.lstTempWoker4Show = new List<PLAN_DR_TEMPWORK>();
            if (null != Request["txtTempWorkSupplierId"])
            {
                string[] aryTempSupplier = Request["txtTempWorkSupplierId"].Split(',');
                string[] aryTempWorkerQty = Request["txtTempWorkerQty"].Split(',');
                string[] aryTempChargeSupplier = Request["txtChargeSupplierId"].Split(',');
                string[] aryTempWarkRemark = Request["txtTempWorkRemark"].Split(',');
                string[] aryDoc = Request["doc"].Split(',');
                for (int i = 0; i < aryTempSupplier.Length; i++)
                {
                    PLAN_DR_TEMPWORK item = new PLAN_DR_TEMPWORK();
                    item.REPORT_ID = RptHeader.REPORT_ID;
                    item.SUPPLIER_ID = aryTempSupplier[i];
                    item.CHARGE_ID = aryTempChargeSupplier[i];
                    ///處理工人名單檔案
                    log.Debug("File Count=" + Request.Files.Count);
                    if (Request.Files[i].ContentLength > 0)
                    {
                        //存檔路徑
                        var fileName = Path.GetFileName(Request.Files[i].FileName);
                        string reportFolder = ContextService.strUploadPath + "\\" + projectid + "\\DailyReport\\";
                        //check 資料夾是否存在
                        string folder = reportFolder + RptHeader.REPORT_ID;
                        ZipFileCreator.CreateDirectory(folder);
                        var path = Path.Combine(folder, fileName);
                        Request.Files[i].SaveAs(path);
                        item.DOC = "DailyReport\\" + RptHeader.REPORT_ID + "\\" + fileName;
                        log.Debug("Upload Sign List File:" + Request.Files[i].FileName);
                    }
                    else
                    {
                        item.DOC = aryDoc[i];
                        log.Error("Not Upload Sign List File!Exist File=" + item.DOC);
                    }
                    log.Debug("Supplier Info=" + item.SUPPLIER_ID);
                    if ("" != aryTempWorkerQty[i].Trim())
                    {
                        item.WORKER_QTY = decimal.Parse(aryTempWorkerQty[i]);
                        newDailyRpt.lstTempWoker4Show.Add(item);
                    }
                    item.REMARK = aryTempWarkRemark[i];
                }
            }
            //處理機具資料
            newDailyRpt.lstRptWorkerAndMachine = new List<PLAN_DR_WORKER>();
            string[] aryMachineType = f["MachineKeyid"].Split(',');
            string[] aryMachineQty = f["planMachineQty"].Split(',');
            for (int i = 0; i < aryMachineType.Length; i++)
            {
                PLAN_DR_WORKER item = new PLAN_DR_WORKER();
                item.REPORT_ID = RptHeader.REPORT_ID;
                item.WORKER_TYPE = "MACHINE";
                item.PARA_KEY_ID = aryMachineType[i];
                if ("" != aryMachineQty[i])
                {
                    item.WORKER_QTY = decimal.Parse(aryMachineQty[i]);
                    newDailyRpt.lstRptWorkerAndMachine.Add(item);
                }
            }
            log.Debug("count MachineD=" + f["MachineKeyid"] + ",WorkerQty=" + f["planMachineQty"]);
            //處理重要事項資料
            newDailyRpt.lstRptNote = new List<PLAN_DR_NOTE>();
            string[] aryNote = f["planNote"].Split(',');
            for (int i = 0; i < aryNote.Length; i++)
            {
                PLAN_DR_NOTE item = new PLAN_DR_NOTE();
                item.REPORT_ID = RptHeader.REPORT_ID;
                if ("" != aryNote[i].Trim())
                {
                    item.SORT = i + 1;
                    item.REMARK = aryNote[i].Trim();
                    newDailyRpt.lstRptNote.Add(item);
                }
            }
            //註記任務是否完成
            if (null == Request["taskDone"])
            {
                newDailyRpt.isDoneFlag = false;
            }
            else
            {
                newDailyRpt.isDoneFlag = true;
            }
            log.Debug("count Note=" + f["planNote"]);
            string msg = planService.createDailyReport(newDailyRpt);
            Response.Redirect("~/ProjectPlan/dailyReport/" + projectid);
            //ProjectPlan/dailyReport/P00061
        }
        //顯示日報維護畫面
        public ActionResult dailyReportList(string id)
        {
            if (null == id || "" == id)
            {
                id = Request["projectid"];
            }
            ViewBag.projectName = planService.getProject(id).PROJECT_NAME;
            ViewBag.projectId = id;
            return View();
        }
        //取得日報明細資料
        public ActionResult getDailyReportList(FormCollection f)
        {
            //定義查詢條件
            string strProjectid = f["txtProjectId"];
            DateTime dtStart = DateTime.MinValue;
            DateTime dtEnd = DateTime.MinValue;
            string strSummary = null;
            if ("" != f["reportDateStart"])
            {
                dtStart = DateTime.Parse(f["reportDateStart"]);
            }
            if ("" != f["reportDateEnd"])
            {
                dtEnd = DateTime.Parse(f["reportDateEnd"]);
            }

            if ("" != f["strSummary"])
            {
                strSummary = f["txtSummary"];
            }
            List<PLAN_DALIY_REPORT> lst = planService.getDailyReportList(strProjectid, dtStart, dtEnd, strSummary);
            ViewBag.Result = "共" + lst.Count + " 筆日報紀錄!!";
            return PartialView("_getDailyReportList", lst);
        }
        //列印施工日報
        public ActionResult dailyReportPrinter()
        {
            DailyReport dailyRpt = null;
            string rptId = Request["rptID"];
            log.Debug("printer report rptid=" + rptId);

            ViewBag.RptId = rptId;
            dailyRpt = planService.getDailyReport(rptId);
            dailyRpt.project = planService.getProject(dailyRpt.dailyRpt.PROJECT_ID);

            //ViewBag.projectId = dailyRpt.dailyRpt.PROJECT_ID;
            //ViewBag.projectName = planService.getProject(dailyRpt.dailyRpt.PROJECT_ID).PROJECT_NAME;
            ViewBag.prj_uid = dailyRpt.lstRptTask[0].PRJ_UID;
            ViewBag.taskName = planService.getProjectTask(dailyRpt.dailyRpt.PROJECT_ID, int.Parse(dailyRpt.lstRptTask[0].PRJ_UID.ToString())).TASK_NAME;

            ViewBag.RptDate = string.Format("{0:yyyy/MM/dd ddd}", dailyRpt.dailyRpt.REPORT_DATE);
            ViewBag.ddlWeather = getDropdownList4Weather(dailyRpt.dailyRpt.WEATHER);

            return View(dailyRpt);
        }
        public ActionResult dailyReport4Estimation(string id)
        {
            if (null == id || "" == id)
            {
                id = Request["projectid"];
            }
            ViewBag.projectName = planService.getProject(id).PROJECT_NAME;
            ViewBag.projectId = id;
            return View();
        }
        /// <summary>
        /// 取得彙整日報表
        /// </summary>
        /// <param name="f"></param>
        /// <returns></returns>
        public ActionResult getDailyReport4Estimation(FormCollection f)
        {
            //定義查詢條件
            string strProjectid = f["txtProjectId"];
            DateTime dtStart= DateTime.Parse(f["reportDateStart"]);
            DateTime dtEnd = DateTime.Parse(f["reportDateEnd"]);

            List<SummaryDailyReport> lst = planService.getDailyReport4Estimation(strProjectid, dtStart, dtEnd);
            Session["dailyReport"] = lst;
            ViewBag.Result = "共" + lst.Count + " 筆估驗紀錄!!";
            return PartialView("_getDailyReport4Estimation", lst);
        }
        /// <summary>
        /// 下載彙整日報表
        /// </summary>
        public void downloadSummaryDailyReport()
        {
            List<SummaryDailyReport> lst = (List<SummaryDailyReport>)Session["dailyReport"];
            SummaryDailyReportToExcel poiservice = new SummaryDailyReportToExcel();
            //產生檔案位置       
            string fileLocation = poiservice.exportExcel(lst);
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
        //彙總報表
        public ActionResult summaryReport()
        {
            List<SummaryDailyReport> lst = null;
            string projectid = Request["projectid"];
            log.Debug("summary report projectid=" + projectid);
            lst = planService.getSummaryReport(projectid);
            return View(lst);
        }
    }
}
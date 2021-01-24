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

namespace topmeperp.Controllers
{
    public class TenderController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        // GET: Tender
        [topmeperp.Filter.AuthFilter]
        public ActionResult Index()
        {
            List<topmeperp.Models.TND_PROJECT> lstProject = SearchProjectByName("", "備標");
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            return View(lstProject);
        }
        // POST : Search
        public ActionResult Search()
        {
            List<topmeperp.Models.TND_PROJECT> lstProject = SearchProjectByName(Request["textProejctName"], Request["selProjectStatus"]);
            ViewBag.SearchResult = "共取得" + lstProject.Count + "筆資料";
            return View("Index", lstProject);
        }
        //POST:Create
        public ActionResult Create(String id)
        {
            logger.Info("get project for update:project_id=" + id);
            TND_PROJECT p = null;
            if (null != id)
            {
                TnderProjectService service = new TnderProjectService();
                p = service.getProjectById(id);
                ViewBag.projectId = p.PROJECT_ID;
            }
            return View(p);
        }
        [HttpPost]
        public ActionResult Create(TND_PROJECT prj, HttpPostedFileBase file)
        {
            logger.Info("create project process! project =" + prj.ToString());
            TnderProjectService service = new TnderProjectService();
            SYS_USER u = (SYS_USER)Session["user"];
            string message = "";
            //1.更新或新增專案基本資料
            if (prj.PROJECT_ID == "" || prj.PROJECT_ID == null)
            {
                //新增專案
                prj.CREATE_USER_ID = u.USER_ID;
                prj.OWNER_USER_ID = u.USER_ID;
                prj.CREATE_DATE = DateTime.Now;
                service.newProject(prj);
                message = "建立專案:" + service.project.PROJECT_ID + "<br/>";
            }
            else
            {
                //修改專案基本資料
                prj.MODIFY_USER_ID = u.USER_ID;
                prj.MODIFY_DATE = DateTime.Now;
                service.updateProject(prj);
                message = "專案基本資料修改:" + prj.PROJECT_ID + "<br/>";
            }
            //若使用者有上傳標單資料，則增加標單資料
            if (null != file && file.ContentLength != 0)
            {
                //2.解析Excel
                logger.Info("Parser Excel data:" + file.FileName);
                //2.1 設定Excel 檔案名稱
                prj.EXCEL_FILE_NAME = file.FileName;
                //2.2 將上傳檔案存檔
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + prj.PROJECT_ID, fileName);
                logger.Info("save excel file:" + path);
                file.SaveAs(path);
                try
                {
                    //2.2 解析Excel 檔案
                    ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                    poiservice.InitializeWorkbook(path);
                    poiservice.ConvertDataForTenderProject(prj.PROJECT_ID, (int)prj.START_ROW_NO);

                    //2.3 記錄錯誤訊息
                    message = message + "標單品項:共" + poiservice.lstProjectItem.Count + "筆資料，";
                    message = message + "<a target=\"_blank\" href=\"/Tender/ManageProjectItem?id=" + prj.PROJECT_ID + "\"> 標單明細檢視畫面單</a><br/>" + poiservice.errorMessage;
                    //        < button type = "button" class="btn btn-primary" onclick="location.href='@Url.Action("ManageProjectItem","Tender", new { id = @Model.tndProject.PROJECT_ID})'; ">標單明細</button>
                    //2.4
                    logger.Info("Delete TND_PROJECT_ITEM By Project ID");
                    service.delAllItemByProject();
                    //2.5
                    logger.Info("Add All TND_PROJECT_ITEM to DB");
                    service.refreshProjectItem(poiservice.lstProjectItem);
                }
                catch (Exception ex)
                {
                    logger.Error("Error Message:" + ex.Message + "." + ex.StackTrace);
                    message = ex.Message;
                }
            }
            ViewBag.projectId = service.project.PROJECT_ID;
            ViewBag.result = message;
            return View(service.project);
        }

        public ActionResult Task(string id)
        {
            logger.Info("task assign page!!");
            ViewBag.projectid = id;
            TnderProjectService service = new TnderProjectService();
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            SYS_USER u = (SYS_USER)Session["user"];
            ViewBag.createid = u.USER_ID;
            return View();
        }

        //POST:TaskAssign
        [HttpPost]
        public ActionResult Task(string id, List<TND_TASKASSIGN> TaskDatas)
        {
            logger.Info("task :" + Request["TaskDatas.index"]);
            TnderProjectService service = new TnderProjectService();
            service.refreshTask(TaskDatas);
            ViewBag.projectid = id;
            return View();
        }
        //編輯任務分派資料
        #region 任務分派
        public ActionResult EditForTask(string id)
        {
            logger.Info("get taskassign for update:task_id=" + id);
            TND_TASKASSIGN ta = null;
            if (null != id)
            {
                TnderProjectService service = new TnderProjectService();
                ta = service.getTaskById(id);
            }
            return View(ta);
        }
        [HttpPost]
        public ActionResult EditForTask(TND_TASKASSIGN taskassign)
        {
            logger.Info("update task assign:task_id=" + taskassign.TASK_ID);
            string message = "";
            TnderProjectService service = new TnderProjectService();
            SYS_USER u = (SYS_USER)Session["user"];
            taskassign.MODIFY_ID = u.USER_ID;
            taskassign.MODIFY_DATE = DateTime.Now;
            if (taskassign.FINISH_DATE.ToString() == "")
            {
                taskassign.FINISH_DATE = DateTime.Now;
            }
            //當任務類型無輸入值時，將View返回頁面同時帶入錯誤訊息
            if (taskassign.TASK_TYPE == null)
            {
                logger.Info("updates fail");
                ViewBag.ErrorMessage = "修改不成功，請重新輸入任務類型，此欄位不可為空白!!";
                return View(taskassign);
            }
            service.updateTask(taskassign);
            message = "任務分派資料修改成功，任務類型為 " + taskassign.TASK_TYPE;
            ViewBag.result = message;
            return View(taskassign);
        }
        #endregion

        public ActionResult Details(string id)
        {
            logger.Info("project detail page projectid = " + id);
            ViewBag.projectid = id;
            List<TND_TASKASSIGN> lstTask = null;
            TnderProjectService service = new TnderProjectService();
            lstTask = service.getTaskByPrjId(id,null);
            TND_PROJECT p = service.getProjectById(id);
            TndProjectModels viewModel = new TndProjectModels();
            var priId = service.getTaskAssignById(id);
            ViewBag.taskAssign = priId;
            viewModel.tndProject = p;
            if (null != priId)
            {
                viewModel.tndTaskAssign = lstTask;
            }
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
        //將專案狀態調整為結案
        public string closeProject()
        {
            string projectid = Request["projectid"];
            string status= Request["status"];
            TnderProjectService service = new TnderProjectService();
            service.closeProject(projectid, status);
            return "狀態已變更!!";
        }
        public ActionResult uploadMapInfo(string id)
        {
            logger.Info("upload map info for projectid=" + id);
            ViewBag.projectid = id;
            return View();
        }
        [HttpPost]
        public ActionResult uploadMapInfo(HttpPostedFileBase fileDevice, HttpPostedFileBase fileFP,
            HttpPostedFileBase fileFW, HttpPostedFileBase fileLCP,
            HttpPostedFileBase filePEP, HttpPostedFileBase filePLU)
        {
            string projectid = Request["projectid"];
            TnderProjectService service = new TnderProjectService();
            string message = "";
            logger.Info("Upload Map Info for projectid=" + projectid);
            try
            {
                //檔案變數名稱需要與前端畫面對應
                #region 設備清單
                //圖算:設備檔案清單
                if (null != fileDevice && fileDevice.ContentLength != 0)
                {
                    //2.解析Excel
                    logger.Info("Parser Excel data:" + fileDevice.FileName);
                    //2.1 設定Excel 檔案名稱
                    var fileName = Path.GetFileName(fileDevice.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    logger.Info("save excel file:" + path);
                    fileDevice.SaveAs(path);
                    //2.2 解析Excel 檔案
                    ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                    poiservice.InitializeWorkbook(path);
                    //解析設備圖算數量檔案
                    List<TND_MAP_DEVICE> lstMapDEVICE = poiservice.ConvertDataForMapDEVICE(projectid);
                    //2.3 記錄錯誤訊息
                    message = message + poiservice.errorMessage;
                    //2.4
                    logger.Info("Delete TND_MAP_DEVICE By Project ID");
                    service.delMapDEVICEByProject(projectid);
                    //2.5
                    logger.Info("Add All TND_MAP_DEVICE to DB");
                    service.refreshMapDEVICE(lstMapDEVICE);
                }
                #endregion
                #region 消防電
                //圖算:消防電(TND_MAP_FP)
                if (null != fileFP && fileFP.ContentLength != 0)
                {
                    //2.解析Excel
                    logger.Info("Parser FP Excel data:" + fileFP.FileName);
                    //2.1 設定Excel 檔案名稱
                    var fileName = Path.GetFileName(fileFP.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    logger.Info("save excel file:" + path);
                    fileFP.SaveAs(path);
                    //2.2 解析Excel 檔案
                    ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                    poiservice.InitializeWorkbook(path);
                    //解析消防電圖算數量檔案
                    List<TND_MAP_FP> lstMapFP = poiservice.ConvertDataForMapFP(projectid);
                    //2.3 記錄錯誤訊息
                    message = message + poiservice.errorMessage;
                    //2.4
                    logger.Info("Delete TND_MAP_FP By Project ID");
                    service.delMapFPByProject(projectid);
                    //2.5
                    logger.Info("Add All TND_MAP_FP to DB");
                    service.refreshMapFP(lstMapFP);
                }
                #endregion
                #region 消防水
                //圖算:消防水(TND_MAP_FW)
                if (null != fileFW && fileFW.ContentLength != 0)
                {
                    //2.解析Excel
                    logger.Info("Parser Excel data:" + fileFW.FileName);
                    //2.1 設定Excel 檔案名稱
                    var fileName = Path.GetFileName(fileFW.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    logger.Info("save excel file:" + path);
                    fileFW.SaveAs(path);
                    //2.2 開啟Excel 檔案
                    ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                    poiservice.InitializeWorkbook(path);
                    //解析消防水塗算數量
                    List<TND_MAP_FW> lstMapFW = poiservice.ConvertDataForMapFW(projectid);
                    //2.3 記錄錯誤訊息
                    message = poiservice.errorMessage;
                    //2.4
                    logger.Info("Delete TND_MAP_FW By Project ID");
                    service.delMapFWByProject(projectid);
                    message = message + "<br/>舊有資料刪除成功 !!";
                    //2.5 
                    logger.Info("Add All TND_MAP_FP to DB");
                    service.refreshMapFW(lstMapFW);
                    message = message + "<br/>資料匯入完成 !!";
                }
                #endregion
                #region 給排水
                //圖算:給排水(TND_MAP_PLU)
                if (null != filePLU && filePLU.ContentLength != 0)
                {
                    //2.解析Excel
                    logger.Info("Parser Excel data:" + filePLU.FileName);
                    //2.1 設定Excel 檔案名稱
                    var fileName = Path.GetFileName(filePLU.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    logger.Info("save excel file:" + path);
                    filePLU.SaveAs(path);
                    //2.2 開啟Excel 檔案
                    ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                    poiservice.InitializeWorkbook(path);
                    //解析給排水圖算數量
                    List<TND_MAP_PLU> lstMapPLU = poiservice.ConvertDataForMapPLU(projectid);
                    //2.3 記錄錯誤訊息
                    message = poiservice.errorMessage;
                    //2.4
                    logger.Info("Delete TND_MAP_PLU By Project ID");
                    service.delMapPLUByProject(projectid);
                    message = message + "<br/>舊有資料刪除成功 !!";
                    //2.5 
                    logger.Info("Add All TND_MAP_PLU to DB");
                    service.refreshMapPLU(lstMapPLU);
                    message = message + "<br/>資料匯入完成 !!";
                }
                #endregion
                #region 弱電管線
                //圖算:弱電管線(TND_MAP_LCP)
                if (null != fileLCP && fileLCP.ContentLength != 0)
                {
                    //2.解析Excel
                    logger.Info("Parser Excel data:" + fileLCP.FileName);
                    //2.1 設定Excel 檔案名稱
                    var fileName = Path.GetFileName(fileLCP.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    logger.Info("save excel file:" + path);
                    fileLCP.SaveAs(path);
                    //2.2 開啟Excel 檔案
                    ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                    poiservice.InitializeWorkbook(path);
                    //解析弱電管線圖算數量
                    List<TND_MAP_LCP> lstMapLCP = poiservice.ConvertDataForMapLCP(projectid);
                    //2.3 記錄錯誤訊息
                    message = poiservice.errorMessage;
                    //2.4
                    logger.Info("Delete TND_MAP_LCP By Project ID");
                    service.delMapLCPByProject(projectid);
                    message = message + "<br/>舊有資料刪除成功 !!";
                    //2.5 
                    logger.Info("Add All TND_MAP_LCP to DB");
                    service.refreshMapLCP(lstMapLCP);
                    message = message + "<br/>資料匯入完成 !!";
                }
                #endregion
                #region 電氣管線
                //圖算:電氣管線(TND_MAP_PEP)
                if (null != filePEP && filePEP.ContentLength != 0)
                {
                    //2.解析Excel
                    logger.Info("Parser Excel data:" + filePEP.FileName);
                    //2.1 設定Excel 檔案名稱
                    var fileName = Path.GetFileName(filePEP.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    logger.Info("save excel file:" + path);
                    filePEP.SaveAs(path);
                    //2.2 開啟Excel 檔案
                    ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                    poiservice.InitializeWorkbook(path);
                    //解析電氣管線圖算數量
                    List<TND_MAP_PEP> lstMapPEP = poiservice.ConvertDataForMapPEP(projectid);
                    //2.3 記錄錯誤訊息
                    message = poiservice.errorMessage;
                    //2.4
                    logger.Info("Delete TND_MAP_PEP By Project ID");
                    service.delMapPEPByProject(projectid);
                    message = message + "<br/>舊有資料刪除成功 !!";
                    //2.5 
                    logger.Info("Add All TND_MAP_PEP to DB");
                    service.refreshMapPEP(lstMapPEP);
                    message = message + "<br/>資料匯入完成 !!";
                }
                #endregion
            }
            catch (Exception ex)
            {
                logger.Error(ex.StackTrace);
                message = ex.Message;
            }
            ViewBag.result = message;
            TempData["result"] = message;
            return RedirectToAction("MapInfoMainPage/" + projectid);

        }
        //建立各類別圖算數量檢視頁面
        public ActionResult MapInfoMainPage(string id)
        {
            logger.Info("mapinfo by projectID=" + id);
            ViewBag.projectid = id;
            string projectid = Request["projectid"];
            List<MAP_FP_VIEW> lstFP = null;
            List<TND_MAP_FW> lstFW = null;
            List<MAP_PEP_VIEW> lstPEP = null;
            List<TND_MAP_PLU> lstPLU = null;
            List<MAP_LCP_VIEW> lstLCP = null;
            List<TND_MAP_DEVICE> lstDEVICE = null;
            if (null != id && id != "")
            {
                TnderProjectService service = new TnderProjectService();
                lstFP = service.getMapFPById(id);
                lstFW = service.getMapFWById(id);
                lstPEP = service.getMapPEPById(id);
                lstPLU = service.getMapPLUById(id);
                lstLCP = service.getMapLCPById(id);
                lstDEVICE = service.getMapDEVICEById(id);
                MapInfoModels viewModel = new MapInfoModels();
                viewModel.mapFP = lstFP;
                viewModel.mapFW = lstFW;
                viewModel.mapPEP = lstPEP;
                viewModel.mapPLU = lstPLU;
                viewModel.mapLCP = lstLCP;
                viewModel.mapDEVICE = lstDEVICE;
                ViewBag.Result1 = "共有" + lstDEVICE.Count + "筆資料";
                ViewBag.Result2 = "共有" + lstFP.Count + "筆資料";
                ViewBag.Result3 = "共有" + lstFW.Count + "筆資料";
                ViewBag.Result4 = "共有" + lstPEP.Count + "筆資料";
                ViewBag.Result5 = "共有" + lstLCP.Count + "筆資料";
                ViewBag.Result6 = "共有" + lstPLU.Count + "筆資料";
                return View(viewModel);
            }
            return RedirectToAction("MapInfoMainPage/" + projectid);
        }
        //編輯消防電圖算數量
        #region 消防電
        public ActionResult EditForFP(string id)
        {
            logger.Info("get map fp for update:fp_id=" + id);
            TND_MAP_FP fp = null;
            if (null != id)
            {
                TnderProjectService service = new TnderProjectService();
                fp = service.getFPById(id);
            }
            return View(fp);
        }
        [HttpPost]
        public ActionResult EditForFP(TND_MAP_FP mapfp)
        {
            logger.Info("update map fp:fp_id=" + mapfp.FP_ID + "," + mapfp.EXCEL_ITEM);
            string message = "";
            TnderProjectService service = new TnderProjectService();
            service.updateMapFP(mapfp);
            message = "消防電圖算資料修改成功，ID :" + mapfp.FP_ID;
            ViewBag.result = message;
            return View(mapfp);
        }
        #endregion
        //編輯消防水圖算數量
        #region 消防水
        public ActionResult EditForFW(string id)
        {
            logger.Info("get map fw for update:fp_id=" + id);
            TND_MAP_FW fw = null;
            if (null != id)
            {
                TnderProjectService service = new TnderProjectService();
                fw = service.getFWById(id);
            }
            return View(fw);
        }
        [HttpPost]
        public ActionResult EditForFW(TND_MAP_FW mapfw)
        {
            logger.Info("update map fw:fw_id=" + mapfw.FW_ID + "," + mapfw.EXCEL_ITEM);
            string message = "";
            TnderProjectService service = new TnderProjectService();
            service.updateMapFW(mapfw);
            message = "消防水圖算資料修改成功，ID :" + mapfw.FW_ID;
            ViewBag.result = message;
            return View(mapfw);
        }
        #endregion
        //編輯電氣管線圖算數量
        #region 電氣管線
        public ActionResult EditForPEP(string id)
        {
            logger.Info("get mappep for update:pep_id=" + id);
            TND_MAP_PEP pep = null;
            if (null != id)
            {
                TnderProjectService service = new TnderProjectService();
                pep = service.getPEPById(id);
            }
            return View(pep);
        }
        [HttpPost]
        public ActionResult EditForPEP(TND_MAP_PEP mappep)
        {
            logger.Info("update map pep:pep_id=" + mappep.PEP_ID + "," + mappep.EXCEL_ITEM);
            string message = "";
            TnderProjectService service = new TnderProjectService();
            service.updateMapPEP(mappep);
            message = "電氣管線圖算資料修改成功，ID :" + mappep.PEP_ID;
            ViewBag.result = message;
            return View(mappep);
        }
        #endregion
        //編輯弱電管線圖算數量
        #region 弱電管線
        public ActionResult EditForLCP(string id)
        {
            logger.Info("get maplcp for update:lcp_id=" + id);
            TND_MAP_LCP lcp = null;
            if (null != id)
            {
                TnderProjectService service = new TnderProjectService();
                lcp = service.getLCPById(id);
            }
            return View(lcp);
        }
        [HttpPost]
        public ActionResult EditForLCP(TND_MAP_LCP maplcp)
        {
            logger.Info("update map lcp:lcp_id=" + maplcp.LCP_ID + "," + maplcp.EXCEL_ITEM);
            string message = "";
            TnderProjectService service = new TnderProjectService();
            service.updateMapLCP(maplcp);
            message = "弱電管線圖算資料修改成功，ID :" + maplcp.LCP_ID;
            ViewBag.result = message;
            return View(maplcp);
        }
        #endregion
        //編輯給排水圖算數量
        #region 給排水
        public ActionResult EditForPLU(string id)
        {
            logger.Info("get mapplu for update:plu_id=" + id);
            TND_MAP_PLU plu = null;
            if (null != id)
            {
                TnderProjectService service = new TnderProjectService();
                plu = service.getPLUById(id);
            }
            return View(plu);
        }
        [HttpPost]
        public ActionResult EditForPLU(TND_MAP_PLU mapplu)
        {
            logger.Info("update map plu:plu_id=" + mapplu.PLU_ID + "," + mapplu.EXCEL_ITEM);
            string message = "";
            TnderProjectService service = new TnderProjectService();
            service.updateMapPLU(mapplu);
            message = "給排水圖算資料修改成功，ID :" + mapplu.PLU_ID;
            ViewBag.result = message;
            return View(mapplu);
        }
        #endregion
        private List<topmeperp.Models.TND_PROJECT> SearchProjectByName(string projectname, string status)
        {
            if (projectname != null)
            {
                logger.Info("search project by 名稱 =" + projectname);
                List<topmeperp.Models.TND_PROJECT> lstProject = new List<TND_PROJECT>();
                using (var context = new topmepEntities())
                {
                    lstProject = context.TND_PROJECT.SqlQuery("select * from TND_PROJECT p "
                        + "where p.PROJECT_NAME Like '%' + @projectname + '%' AND ISNULL(STATUS,'備標')=@status;",
                         new SqlParameter("projectname", projectname), new SqlParameter("status", status)).ToList();
                }
                logger.Info("get project count=" + lstProject.Count);
                return lstProject;
            }
            else
            {
                return null;
            }
        }
        /// <summary>
        /// 設定標單品項查詢條件
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public ActionResult ManageProjectItem(string id)
        {
            //傳入專案編號，
            InquiryFormService service = new InquiryFormService();
            logger.Info("start project id=" + id);

            //取得專案基本資料fc
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.id = p.PROJECT_ID;
            ViewBag.projectName = p.PROJECT_NAME;

            SelectListItem empty = new SelectListItem();
            empty.Value = "";
            empty.Text = "";
            List<SelectListItem> selectMain = UtilService.getMainSystem(id, service);
            // selectMain.Add(empty);
            ViewBag.SystemMain = selectMain;
            List<SelectListItem> selectSub = UtilService.getSubSystem(id, service);
            //selectSub.Add(empty);
            ViewBag.SystemSub = selectSub;
            //設定查詢條件
            return View();
        }

        /// <summary>
        /// 取得標單明細資料
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        public ActionResult ShowProejctItems(FormCollection form)
        {
            InquiryFormService service = new InquiryFormService();
            logger.Info("start project id=" + Request["id"] + ",TypeCode1=" + Request["typeCode1"] + ",typecode2=" + Request["typeCode2"] + ",SystemMain=" + Request["SystemMain"] + ",Sytem Sub=" + Request["SystemSub"]);
            logger.Debug("Exception check=" + Request["chkEx"]);
            List<TND_PROJECT_ITEM> lstItems = service.getProjectItem(Request["chkEx"], Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"], Request["selDelFlag"]);
            //logger.Debug("TEST:使用壓力:≥7kg/c㎡G");
            //foreach (var p in lstItems)
            //{
            //    logger.Debug("PK=" + p.PROJECT_ITEM_ID + ",ITEM_DESC=" + p.ITEM_DESC + ",excel row=" + p.EXCEL_ROW_ID);
            //}

            ViewBag.Result = "共幾" + lstItems.Count + "筆資料";
            return PartialView(lstItems);
        }
        /// <summary>
        /// 取得標單品項詳細資料
        /// </summary>
        /// <param name="itemid"></param>
        /// <returns></returns>
        public string getProjectItem(string itemid)
        {
            InquiryFormService service = new InquiryFormService();
            logger.Info("get project item by id=" + itemid);
            //TND_PROJECT_ITEM  item = service.getProjectItem.getUser(userid);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(service.getProjectItem(itemid));
            logger.Info("project item  info=" + itemJson);
            return itemJson;
        }
        /// <summary>
        /// 新增或更新Project_item 資料
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        public String addProjectItem(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "更新成功!!";

            TND_PROJECT_ITEM item = new TND_PROJECT_ITEM();
            item.PROJECT_ID = form["project_id"];
            item.PROJECT_ITEM_ID = form["project_item_id"];
            item.ITEM_ID = form["item_id"];
            item.ITEM_DESC = form["item_desc"];
            item.ITEM_UNIT = form["item_unit"];
            try
            {
                item.ITEM_QUANTITY = decimal.Parse(form["item_quantity"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.PROJECT_ITEM_ID + " not quattity:" + ex.Message);
            }
            try
            {
                item.ITEM_UNIT_PRICE = decimal.Parse(form["item_unit_price"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.PROJECT_ITEM_ID + " not unit price:" + ex.Message);
            }
            item.ITEM_REMARK = form["item_remark"];
            if (form["type_code_1"].Trim() != "")
            {
                item.TYPE_CODE_1 = form["type_code_1"];
            }
            else
            {
                item.TYPE_CODE_1 = null;
            }

            if (form["type_code_2"].Trim() != "")
            {
                item.TYPE_CODE_2 = form["type_code_2"];
            }
            else
            {
                item.TYPE_CODE_2 = null;
            }

            item.SYSTEM_MAIN = form["system_main"];
            item.SYSTEM_SUB = form["system_sub"];
            item.DEL_FLAG = form["selDelFlag"];
            try
            {
                item.EXCEL_ROW_ID = long.Parse(form["excel_row_id"]);
            }
            catch (Exception ex)
            {
                logger.Error(item.PROJECT_ITEM_ID + " not exce row id:" + ex.Message);
            }

            SYS_USER loginUser = (SYS_USER)Session["user"];
            item.MODIFY_USER_ID = loginUser.USER_ID;
            item.MODIFY_DATE = DateTime.Now;
            InquiryFormService service = new InquiryFormService();
            int i = 0;
            string strFlag = form["flag"].Trim();
            if (strFlag.Equals("addAfter"))
            {
                i = service.addProjectItemAfter(item);
            }
            else
            {
                i = service.updateProjectItem(item);
            }

            if (i == 0) { msg = service.message; }
            return msg;
        }
        /// <summary>
        /// Project_item 註記刪除
        /// </summary>
        /// <param name="itemid"></param>
        /// <returns></returns>
        public String delProjectItem(string itemid)
        {
            InquiryFormService service = new InquiryFormService();
            string msg = "更新成功!!";
            logger.Info("del project item by id=" + itemid);
            int i = service.changeProjectItem(itemid, "Y");
            return msg + "(" + i + ")";
        }
        //上傳得標後標單內容(用於轉換專案狀態)
        [HttpPost]
        public ActionResult uploadPlanItem(TND_PROJECT prj, HttpPostedFileBase file)
        {
            //1.取得專案編號
            string projectid = Request["projectid"];
            logger.Info("Upload plan items for projectid=" + projectid);
            TnderProjectService service = new TnderProjectService();
            SYS_USER u = (SYS_USER)Session["user"];
            string message = "";
            //修改專案狀態
            prj.PROJECT_ID = projectid;
            prj.STATUS = "專案執行";
            prj.PROJECT_NAME = Request["projectname"];
            prj.ENG_NAME = Request["engname"];
            prj.CONTRUCTION_NAME = Request["contructionname"];
            prj.LOCATION = Request["location"];
            prj.OWNER_NAME = Request["ownername"];
            prj.CONTACT_NAME = Request["contactname"];
            prj.CONTACT_TEL = Request["contecttel"];
            prj.CONTACT_FAX = Request["contactfax"];
            prj.CONTACT_EMAIL = Request["contactemail"];
            prj.DUE_DATE = Convert.ToDateTime(Request["duedate"]);
            prj.SCHDL_OFFER_DATE = Convert.ToDateTime(Request["schdlofferdate"]);
            if (Request["starrowno"] != null || " " != Request["starrowno"])
            { prj.START_ROW_NO = int.Parse(Request["starrowno"]); }
            //prj.WAGE_MULTIPLIER = decimal.Parse(Request["wage"]);
            prj.EXCEL_FILE_NAME = Request["excelfilename"];
            prj.OWNER_USER_ID = Request["owneruserid"];
            prj.CREATE_USER_ID = Request["createdid"];
            prj.CREATE_DATE = Convert.ToDateTime(Request["createddate"]);
            prj.MODIFY_USER_ID = u.USER_ID;
            prj.MODIFY_DATE = DateTime.Now;
            service.updateProject(prj);
            message = "專案進入執行階段:" + prj.PROJECT_ID + "<br/>";

            if (null != file && file.ContentLength != 0)
            {
                //2.解析Excel
                logger.Info("Parser Excel data:" + file.FileName);
                //2.1 將上傳檔案存檔
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                logger.Info("save excel file:" + path);
                file.SaveAs(path);
                //2.2 解析Excel 檔案
                ProjectItemFromExcel poiservice = new ProjectItemFromExcel();
                poiservice.InitializeWorkbook(path);
                poiservice.ConvertDataForPlan(projectid);
                //2.3 記錄錯誤訊息
                message = message + "得標標單品項:共" + poiservice.lstPlanItem.Count + "筆資料，";
                message = message + "<a target=\"_blank\" href=\"/Plan/ManagePlanItem?id=" + projectid + "\"> 標單明細檢視畫面單</a><br/>" + poiservice.errorMessage;
                //        < button type = "button" class="btn btn-primary" onclick="location.href='@Url.Action("ManagePlanItem","Plan", new { id = @Model.tndProject.PROJECT_ID})'; ">標單明細</button>
                //2.4
                logger.Info("Delete PLAN_ITEM By Project ID");
                service.delAllItemByPlan();
                //2.5
                logger.Info("Add All PLAN_ITEM to DB");
                service.refreshPlanItem(poiservice.lstPlanItem);
            }
            TempData["result"] = message;//Plan/ManagePlanItem/P00023
            //return RedirectToAction()
            return RedirectToAction("ManagePlanItem/" + projectid, "Plan");
        }

        /// <summary>
        /// 下載現有標單資料
        /// </summary>
        public void downLoadProjectItem()
        {
            string projectid = Request["projectid"];
            ProjectItem2Excel poiservice = new ProjectItem2Excel();
            //產生檔案位置
            poiservice.exportExcel(projectid);
            string fileLocation = "/UploadFile/" + projectid + "/" + projectid + "_標單明細.xlsx";
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
        /// <summary>
        /// 上載現有標單資料
        /// </summary>
        public string uploadProjectItem(HttpPostedFileBase file1)
        {
            string projectid = Request["id"];
            logger.Debug("ProjectID=" + projectid + ",Upload ProjectItem=" + file1.FileName);
            TnderProjectService service = new TnderProjectService();
            service.getProjectById(projectid);
            SYS_USER u = (SYS_USER)Session["user"];

            if (null != file1 && file1.ContentLength != 0)
            {
                try
                {
                    //2.解析Excel
                    logger.Info("Parser Excel data:" + file1.FileName);
                    //2.1 將上傳檔案存檔
                    var fileName = Path.GetFileName(file1.FileName);
                    var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                    logger.Info("save excel file:" + path);
                    file1.SaveAs(path);
                    //2.2 解析Excel 檔案
                    //poiservice.ConvertDataForTenderProject(prj.PROJECT_ID, (int)prj.START_ROW_NO);
                    ProjectItem2Excel poiservice = new ProjectItem2Excel();
                    poiservice.InitializeWorkbook(path);
                    poiservice.ConvertDataForTenderProject(projectid);
                    //2.3 記錄錯誤訊息
                    //        < button type = "button" class="btn btn-primary" onclick="location.href='@Url.Action("ManagePlanItem","Plan", new { id = @Model.tndProject.PROJECT_ID})'; ">標單明細</button>
                    //2.4
                    logger.Info("Delete Project_Item  By Project ID");
                    service.delAllItemByProject();
                    //2.5
                    logger.Info("Add All Project_ITEM to DB");
                    service.refreshProjectItem(poiservice.lstProjectItem);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.StackTrace);
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

        //新增任務分派資料
        public String AddTaskAssign(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "新增任務分派資料成功!!";
            TND_TASKASSIGN leader = new TND_TASKASSIGN();
            TND_TASKASSIGN costing = new TND_TASKASSIGN();
            TND_TASKASSIGN map = new TND_TASKASSIGN();
            leader.PROJECT_ID = form["project_id"];
            costing.PROJECT_ID = form["project_id"];
            map.PROJECT_ID = form["project_id"];
            leader.TASK_TYPE = "主辦";
            costing.TASK_TYPE = "成控";
            map.TASK_TYPE = "圖算";
            if (form["leader_user_id"] != "")
            {
                leader.USER_ID = form["leader_user_id"];
            }
            if (form["leader_task_item"] != "")
            {
                leader.TASK_ITEM = form["leader_task_item"];
            }
            if (form["leader_remark"] != "")
            {
                leader.REMARK = form["leader_remark"];
            }
            if (form["leader_finish_date"] != "")
            {
                leader.FINISH_DATE = Convert.ToDateTime(form.Get("leader_finish_date"));
            }
            if (form["costing_user_id"] != "")
            {
                costing.USER_ID = form["costing_user_id"];
            }
            if (form["costing_task_item"] != "")
            {
                costing.TASK_ITEM = form["costing_task_item"];
            }
            if (form["costing_remark"] != "")
            {
                costing.REMARK = form["costing_remark"];
            }
            if (form["costing_finish_date"] != "")
            {
                costing.FINISH_DATE = Convert.ToDateTime(form.Get("costing_finish_date"));
            }
            if (form["map_user_id"] != "")
            {
                map.USER_ID = form["map_user_id"];
            }
            if (form["map_task_item"] != "")
            {
                map.TASK_ITEM = form["map_task_item"];
            }
            if (form["map_remark"] != "")
            {
                map.REMARK = form["map_remark"];
            }
            if (form["map_finish_date"] != "")
            {
                map.FINISH_DATE = Convert.ToDateTime(form.Get("map_finish_date"));
            }
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            leader.CREATE_ID = uInfo.USER_ID;
            leader.CREATE_DATE = DateTime.Now;
            costing.CREATE_ID = uInfo.USER_ID;
            costing.CREATE_DATE = DateTime.Now;
            map.CREATE_ID = uInfo.USER_ID;
            map.CREATE_DATE = DateTime.Now;
            List<TND_TASKASSIGN> taskAssign = new List<TND_TASKASSIGN>();
            taskAssign.Add(leader);
            taskAssign.Add(costing);
            taskAssign.Add(map);
            TnderProjectService service = new TnderProjectService();
            int i = service.refreshTask(taskAssign);
            if (i == 0) { msg = service.message; }
            return msg;
        }
        /// <summary>
        /// 取得任務分派詳細資料
        /// </summary>
        /// <param name="itemid"></param>
        /// <returns></returns>
        public string getTaskItem(string itemid)
        {
            TnderProjectService service = new TnderProjectService();
            logger.Info("get task assign item by id=" + itemid);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(service.getTaskById(itemid));
            logger.Info("task assign item  info=" + itemJson);
            return itemJson;
        }
        //更新任務分派資料
        public String refreshTaskItem(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "修改任務分派資料成功!!";
            TND_TASKASSIGN item = new TND_TASKASSIGN();
            item.PROJECT_ID = form["prjid"];
            if (null != form["task_id"] && form["task_id"] != "")
            {
                item.TASK_ID = Int64.Parse(form["task_id"]);
            }
            item.USER_ID = form["userId"];
            item.TASK_TYPE = form["taskType"];
            item.TASK_ITEM = form["taskItem"];
            item.REMARK = form["taskRemark"];
            if (null != form["finishDate"] && form["finishDate"] != "")
            {
                item.FINISH_DATE = Convert.ToDateTime(form.Get("finishDate"));
            }

            item.CREATE_DATE = DateTime.Now;
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            item.CREATE_ID = u.USER_ID;
            item.MODIFY_ID = u.USER_ID;
            item.MODIFY_DATE = DateTime.Now;
            TnderProjectService service = new TnderProjectService();
            int i = service.updateTask(item);
            if (i == 0) { msg = service.message; }
            return msg;
        }
    }
}
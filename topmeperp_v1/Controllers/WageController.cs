using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;
using System.IO;

namespace topmeperp.Controllers
{
    public class WageController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(WageController));
        WageTableService service = new WageTableService();
        // GET: 下載工率excel表格
        public ActionResult Index(string id, FormCollection form)
        {
            log.Info("get project item :projectid=" + id);
            //取得專案基本資料
            ViewBag.projectid = id;
            service.getProjectId(id);
            WageFormToExcel poi = new WageFormToExcel();
            poi.exportExcel(service.wageTable, service.wageTableItem);
            TND_PROJECT p = service.getProjectById(id);
            ViewBag.projectName = p.PROJECT_NAME;
            return View();
        }
        //上傳工率
        [HttpPost]
        public ActionResult uploadWageTable(HttpPostedFileBase fileWage)
        {

            string projectid = Request["projectid"];
            log.Info("Upload Wage Table for projectid=" + projectid);
            string message = "";
            //檔案變數名稱需要與前端畫面對應
            if (null != fileWage && fileWage.ContentLength != 0)
            {
                //2.解析Excel
                log.Info("Parser Excel data:" + fileWage.FileName);
                //2.1 設定Excel 檔案名稱
                var fileName = Path.GetFileName(fileWage.FileName);
                var path = Path.Combine(ContextService.strUploadPath + "/" + projectid, fileName);
                log.Info("save excel file:" + path);
                fileWage.SaveAs(path);
                //2.2 開啟Excel 檔案
                log.Info("Parser Excel File Begin:" + fileWage.FileName);
                WageFormToExcel wageservice = new WageFormToExcel();
                wageservice.InitializeWorkbook(path);
                //解析工率數量
                List<TND_WAGE> lstWage = wageservice.ConvertDataForWage(projectid);
                //2.3 記錄錯誤訊息
                message = wageservice.errorMessage;
                //2.4
                log.Info("Delete TND_WAGE By Project ID");
                service.delWageByProject(projectid);
                message = message + "<br/>舊有資料刪除成功 !!";
                //2.5 
                log.Info("Add All TND_WAGE to DB");
                service.refreshWage(lstWage);
                message = message + "<br/>資料匯入完成 !!";
            }
            TempData["result"] = message;
            // 將工率乘數寫入工率Table的Price欄位
            int k = service.updateWagePrice(projectid);
            return RedirectToAction("WageTable/" + projectid);
        }

        public ActionResult WageTable(string id)
        {
            log.Info("wage ratio by projectID=" + id);
            ViewBag.projectid = id;
            List<TND_WAGE> lstWage = null;
            if (null != id && id != "")
            {
                WageTableService service = new WageTableService();
                lstWage = service.getWageById(id);
                ViewBag.Result = "共有" + lstWage.Count + "筆資料";
            }
            return View(lstWage);
        }
        
        // GET: Wage/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: Wage/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: Wage/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: Wage/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}

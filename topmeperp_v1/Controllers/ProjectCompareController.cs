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
    public class ProjectCompareController : Controller
    {
        RptCompareProjectPrice service = new RptCompareProjectPrice();
        static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        // GET: ProjectCompare
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult getCompareData(FormCollection f)
        {
            log.Info("Source ProijectID=" + f["srcprojectid"] + ",Target ProjectId=" + f["tarprojectid"] + "," + f["hasPrice"]);
            bool hasPriec = false;
            bool hasProject = false;
            if (null != f["hasPrice"])
            {
                hasPriec = true;
            }
            if (null != f["hasProject"])
            {
                hasProject = true;
            }

            List<ProjectCompareData> lst = service.RtpGetPriceFromExistProject(f["srcprojectid"], f["tarprojectid"], hasProject, hasPriec);
            ViewBag.Result = "共取得" + lst.Count + "筆資料!!";
            return PartialView("_CompareData", lst);
        }
        public ActionResult Update(FormCollection f)
        {
            //item.SOURCE_PROJECT_ID + '|' + item.SOURCE_SYSTEM_MAIN + '|' + item.SOURCE_SYSTEM_SUB + '|' + item.SRC_UNIT_PRICE + '|' + item.TARGET_PROJECT_ID + '|' + item.SOURCE_ITEM_DESC;}
            string[] lstItem = f["chkItem"].Split(',');
            List<ProjectCompareData> lstComparedata = new List<ProjectCompareData>();
            for (int i = 0; i < lstItem.Count(); i++)
            {
                log.Debug("ITEM_INFO=" + lstItem[i]);
                string[] data = lstItem[i].Split('|');
                ProjectCompareData item = new ProjectCompareData();
                item.SOURCE_PROJECT_ID = data[0];
                item.SOURCE_SYSTEM_MAIN = data[1];
                item.SOURCE_SYSTEM_SUB = data[2];
                if (data[3] != "")
                {
                    item.SRC_UNIT_PRICE = decimal.Parse(data[3]);
                }
                item.TARGET_PROJECT_ID = data[4];
                item.TARGET_ITEM_DESC = data[5].Replace("xyz", ",");
                lstComparedata.Add(item);
            }
            int j = service.MigratePrice(lstComparedata);
            log.Info("更新" + j + "筆");
            return View("Index");
        }
    }
}
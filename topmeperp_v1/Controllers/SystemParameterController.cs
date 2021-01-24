using log4net;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;

namespace topmeperp.Controllers
{
    public class SystemParameterController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(UserManageController));
        SystemParameter service = new SystemParameter();
        // GET: SystemParameter
        /// <summary>
        /// 取得功能參數清單
        /// </summary>
        /// <returns></returns>
        public ActionResult Index()
        {
            log.Info("index");
            SelectList paraFunction = new SelectList(service.getParaIndex(), "FUNCTION_ID", "FUNCTION_DESC");
            ViewBag.paraFunction = paraFunction;
            return View();
        }
        /// <summary>
        /// 取得欄位參數清單
        /// </summary>
        /// <returns></returns>
        public JsonResult GetFields(string functionId)
        {
            log.Info("functionId="+ functionId);
            List<KeyValuePair<string, string>> items = new List<KeyValuePair<string, string>>();
            if (!string.IsNullOrWhiteSpace(functionId))
            {
                var fields = service.getFieldIndex(functionId);
                if (fields.Count() > 0)
                {
                    foreach (var field in fields)
                    {
                        items.Add(new KeyValuePair<string, string>(field.FUNCTION_ID,field.FUNCTION_DESC));
                    }
                }
            }
            log.Debug("json:" + this.Json(items));
            return this.Json(items);
        }
        /// <summary>
        /// 取得Key ,Value
        /// </summary>
        /// <returns></returns>
        public PartialViewResult getKeyValues()
        {
            string functionId = Request["paraFunction"];
            string fieldId = Request["paraField"];
            List<SYS_PARA>  lst = null;
            if (fieldId != "")
            {
                lst = SystemParameter.getSystemPara(functionId, fieldId);
            }
            else
            {
                lst = SystemParameter.getSystemPara(functionId);
            }
            log.Debug("functionId,fieldId" + functionId + "," + fieldId);
            return PartialView("_Para", lst) ;
        }
        public string updateSysPara()
        {
            string[] lstParaId = Request["paraId"].Split(',');
            string[] lstKey = Request["keyField"].Split(',');
            string[] lstValue = Request["valueField"].Split(',');
            List<SYS_PARA> lst = new List<SYS_PARA>();
            for (int i = 0; i < lstParaId.Length; i++)
            {
                SYS_PARA p = new SYS_PARA();
                p.PARA_ID =Convert.ToInt32(lstParaId[i]);
                p.VALUE_FIELD = lstValue[i];
                p.KEY_FIELD = lstKey[i];
                lst.Add(p);
            }
            string json = JsonConvert.SerializeObject(lst);
            log.Debug("json=" + json);
            string result= SystemParameter.UpdateSysPara(lst);
            return result;
        }
    }
}
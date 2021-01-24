using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;

namespace topmeperp.Service
{
    public class UtilService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public static List<SelectListItem> getMainSystem(string id, InquiryFormService service)
        {
            //取得主系統資料
            List<SelectListItem> selectMain = new List<SelectListItem>();
            foreach (string itm in service.getSystemMain(id))
            {
                logger.Debug("Main System=" + itm);
                SelectListItem selectI = new SelectListItem();
                selectI.Value = itm;
                selectI.Text = itm;
                if (null != itm && "" != itm)
                {
                    selectMain.Add(selectI);
                }
            }
            return selectMain;
        }
        public static List<SelectListItem> getSubSystem(string id, InquiryFormService service)
        {
            //取得次系統資料
            List<SelectListItem> selectSub = new List<SelectListItem>();
            foreach (string itm in service.getSystemSub(id))
            {
                logger.Debug("Sub System=" + itm);
                SelectListItem selectI = new SelectListItem();
                selectI.Value = itm;
                selectI.Text = itm;
                if (null != itm && "" != itm)
                {
                    selectSub.Add(selectI);
                }
            }

            return selectSub;
        }
        public static string covertToJson(Object obj)
        {
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(obj);
            logger.Debug(itemJson);
            return itemJson;
        }
        public static SYS_USER getUserInfoFromSession( HttpSessionStateBase Session )
        {
            SYS_USER u = (SYS_USER)Session["user"];
            return u;
        }
    }
}
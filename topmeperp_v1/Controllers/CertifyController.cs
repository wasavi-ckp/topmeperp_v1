using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Filter;
using topmeperp.Models;
using topmeperp.Service;

namespace topmeperp.Controllers
{
    public class CertifyController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        // GET: Certify
        private CertifyService _db = new CertifyService();
        //protected override void Dispose(bool disposing)                              //冠蒲修訂------解除資料庫連線

        //{
        //    if (disposing)
        //    {
        //        _db.Dispose();
        //    }
        //    base.Dispose(disposing);
        //}
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult getPlanSupInquiryByProjectId(string _id)
        {
            logger.Fatal("getPlanSupInquiryByProjectId : " + _id);
            List<PLAN_SUP_INQUIRY> cS = _db.getPlanSupInquiryByProjectId(_id);
            return View(cS);

        }
        public ActionResult getPlanSupInquiryItemByInquiryFormId(string _id)
        {
            logger.Fatal("getPlanSupInquiryItemByInquiryFormId : " + _id);
            List<PLAN_SUP_INQUIRY_ITEM> cs = _db.getPlanSupInquiryItemByInquiryFormId(_id);
            return View(cs);
        }




    }
}

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
    public class TypeManageController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        TypeManageService service = new TypeManageService();
        // GET: TypeManage
        public ActionResult Index()
        {
            List<REF_TYPE_MAIN> lst = service.getMainType();
            return View(lst);
        }
        public ActionResult getSubType()
        {
            string typecod1 = Request["typecode1"];
            string typecod2 = Request["typecode2"];
            List<REF_TYPE_SUB> lst = service.getSubType(typecod1.Trim() + typecod2.Trim());
            logger.Debug("get main type code1=" + typecod1 + ",typecod2=" + typecod2);

            @ViewBag.SearchResult = "取得" + lst.Count + "筆資料";
            @ViewBag.Typecode = "<a href =\"EditMainType?typecode1=" + typecod1 + "&typecode2=" + typecod2 + "\">編輯</a>";
            return PartialView("_getSubType", lst);
        }
        public ActionResult EditMainType()
        {
            string typecod1 = Request["typecode1"];
            string typecod2 = Request["typecode2"];
            logger.Debug("get main type code1=" + typecod1 + ",typecod2=" + typecod2);
            TyepManageModel typeManageModel = service.getTypeManageModel(typecod1, typecod2);
            return View(typeManageModel);
        }
        [HttpPost]
        public ActionResult EditMainType(TyepManageModel typeManageModel)
        {
            logger.Info("modify mainType" + typeManageModel.MainType.TYPE_DESC);
            logger.Info("modify SubType" + Request["item.SUB_TYPE_CODE"]);
            string[] subTypeId = null;
            string[] mainTypeId = null;
            string[] subTypeCode = null;
            string[] subTypeDesc = null;
            if (Request["item.SUB_TYPE_ID"] != null)
            {
                subTypeId = Request["item.SUB_TYPE_ID"].Split(',');
                mainTypeId = Request["item.TYPE_CODE_ID"].Split(',');
                subTypeCode = Request["item.SUB_TYPE_CODE"].Split(',');
                subTypeDesc = Request["item.TYPE_DESC"].Split(',');
            }
            List<REF_TYPE_SUB> lstSubType = new List<REF_TYPE_SUB>();
            if (subTypeId != null)
            {
                for (int i = 0; i < subTypeCode.Count(); i++)
                {
                    logger.Debug("subTypeId=" + subTypeId[i] + ",mainTypeId=" + mainTypeId[i] + ",subTypeCode=" + subTypeCode[i] + ",subTypeDesc=" + subTypeDesc[i]);
                    REF_TYPE_SUB subType = new REF_TYPE_SUB();
                    if (subTypeId[i] == "")
                    {
                        //新次九宮格
                        subType.SUB_TYPE_ID = int.Parse(mainTypeId[i] + subTypeCode[i]);
                    }
                    else
                    {
                        //原次九宮格
                        subType.SUB_TYPE_ID = int.Parse(subTypeId[i]);
                    }
                    subType.TYPE_CODE_ID = mainTypeId[i];
                    subType.SUB_TYPE_CODE = int.Parse(subTypeCode[i]);
                    subType.TYPE_DESC = subTypeDesc[i];
                    lstSubType.Add(subType);
                }
            }
            service.updateTypeManageModel(typeManageModel.MainType, lstSubType);
            return Redirect("Index");
            //return View(typeManageModel);
        }
    }
}



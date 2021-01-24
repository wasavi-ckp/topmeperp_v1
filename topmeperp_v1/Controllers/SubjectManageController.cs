using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using log4net;
using topmeperp.Service;
using topmeperp.Models;
using System.Data.SqlClient;

namespace topmeperp.Controllers
{
    public class SubjectManageController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        SubjectManageService service = new SubjectManageService();

        // GET: SubjectManage
        public ActionResult Index()
        {
            logger.Info("index");
            //取得財務科目類別
            List<string> finCategory = service.getCategory();
            ViewData.Add("categorys", finCategory);
            SelectList categorys = new SelectList(finCategory, "CATEGORY");
            ViewBag.categorys = categorys;
            //將資料存入TempData 減少不斷讀取資料庫
            TempData.Remove("categorys");
            TempData.Add("categorys", finCategory);
            return View();
        }

        //關鍵字尋找財務科目
        public ActionResult Search()
        {
            List<FIN_SUBJECT> lstSubject = SearchSubjectByName(Request["SubjectName"], Request["categorys"]);
            //取得財務科目類別
            List<string> finCategory = service.getCategory();
            ViewData.Add("categorys", finCategory);
            SelectList categorys = new SelectList(finCategory, "CATEGORY");
            ViewBag.categorys = categorys;
            //將資料存入TempData 減少不斷讀取資料庫
            TempData.Remove("categorys");
            TempData.Add("categorys", finCategory);
            ViewBag.SearchResult = "共取得" + lstSubject.Count + "筆資料";
            return View("Index", lstSubject);
        }
        private List<FIN_SUBJECT> SearchSubjectByName(string subjectname, string category)
        {
            if (subjectname != null)
            {
                logger.Info("search subject by 名稱 =" + subjectname + ", by 項目類別 =" + category);
                List<FIN_SUBJECT> lstSubject = new List<FIN_SUBJECT>();
                string sql = "SELECT * FROM FIN_SUBJECT WHERE FIN_SUBJECT_ID IS NOT NULL  ";

                var parameters = new List<SqlParameter>();
                //九宮格條件
                if (null != category && "" != category)
                {
                    //sql = sql + " AND CATEGORY ='" + category + "'";
                    sql = sql + " AND CATEGORY=@category ";
                    parameters.Add(new SqlParameter("category", category));
                }
                //項目名稱條件
                if (null != subjectname && "" != subjectname)
                {
                    //sql = sql + " AND SUBJECT_NAME LIKE '%' + @subjectname + '%' ;
                    sql = sql + " AND SUBJECT_NAME LIKE @subjectname";
                    parameters.Add(new SqlParameter("subjectname", '%' + @subjectname + '%'));
                }
                sql = sql + " ORDER BY FIN_SUBJECT_ID ";
                logger.Info("sql=" + sql);
                using (var context = new topmepEntities())
                {
                    lstSubject = context.FIN_SUBJECT.SqlQuery(sql, parameters.ToArray()).ToList();

                }
                logger.Info("get subject count=" + lstSubject.Count);
                return lstSubject;
            }
            else
            {
                return null;
            }
        }
        public string getSubjectItem(string itemid)
        {
            logger.Info("get subject item by id=" + itemid);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(service.getSubjectItem(itemid));
            logger.Info("subject item  info=" + itemJson);
            return itemJson;
        }

        //新增或修改財務項目
        public String addSubject(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            string msg = "";
            //懶得把Form綁FIN_SUBJECT 直接先把Form 值填滿
            FIN_SUBJECT s = new FIN_SUBJECT();
            if (null != form.Get("subject_id").Trim() && form.Get("subject_id").Trim() != "")
            {
                s.FIN_SUBJECT_ID = form.Get("subject_id").Trim();
            }
            else
            {
                s.FIN_SUBJECT_ID = form.Get("new_subject_id").Trim();
            }
            s.SUBJECT_NAME = form.Get("subject_name").Trim();
            s.CATEGORY = form.Get("categorys").Trim();
            int i = service.addNewSubject(s);
            if (i == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "項目更新成功";
            }

            logger.Info("Request:SUBJECT_ID=" + s.FIN_SUBJECT_ID);
            return msg;
        }
    }
}

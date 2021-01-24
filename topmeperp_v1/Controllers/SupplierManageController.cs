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
    public class SupplierManageController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(SupplierManageController));

        // GET: SupplierManage
        public ActionResult Index()
        {
            log.Info("index");
            return View();
        }
        //關鍵字尋找供應商
        public ActionResult Search()
        {
            List<TND_SUPPLIER> lstSupplier = SearchSupplierByName(Request["textSupplierName"], Request["textSuppyNote"], Request["textTypeMain"]);
            ViewBag.SearchResult = "共取得" + lstSupplier.Count + "筆資料";
            return View("Index", lstSupplier);
        }

        private List<TND_SUPPLIER> SearchSupplierByName(string suppliername, string supplyNote, string typeMain)
        {
            if (suppliername != null)
            {
                log.Info("search supplier by 名稱 =" + suppliername + ", by 九宮格 =" + typeMain + ", by 產品類別 =" + supplyNote);
                List<TND_SUPPLIER> lstSupplier = new List<TND_SUPPLIER>();
                string sql = "SELECT * FROM TND_SUPPLIER s WHERE s.SUPPLIER_ID IS NOT NULL  ";

                var parameters = new List<SqlParameter>();
                //九宮格條件
                if (null != typeMain && "" != typeMain)
                {
                    //sql = sql + " AND sm.TYPE_MAIN ='" + typeMain + "'";
                    sql = sql + " AND TYPE_MAIN=@typeMain";
                    parameters.Add(new SqlParameter("typeMain", typeMain));
                }
                //供應商名稱條件
                if (null != suppliername && "" != suppliername)
                {
                    //sql = sql + " AND s.COMPANY_NAME LIKE '%' + @suppliername + '%' ;
                    sql = sql + " AND COMPANY_NAME LIKE @suppliername";
                    parameters.Add(new SqlParameter("suppliername", '%' + @suppliername + '%'));
                }
                //產品類別條件
                if (null != supplyNote && "" != supplyNote)
                {
                    //sql = sql + " AND sm.SUPPLY_NOTE LIKE '%' + @supplyNote + '%' ;
                    sql = sql + " AND SUPPLY_NOTE LIKE @supplyNote";
                    parameters.Add(new SqlParameter("supplyNote", '%' + @supplyNote + '%'));
                }
                sql = sql + " ORDER BY TYPE_MAIN ";
                log.Info("sql=" + sql);
                using (var context = new topmepEntities())
                {
                    lstSupplier = context.TND_SUPPLIER.SqlQuery(sql, parameters.ToArray()).ToList();

                }
                log.Info("get supplier count=" + lstSupplier.Count);
                return lstSupplier;
            }
            else
            {
                return null;
            }
        }
        //取得供應商資料
        public ActionResult Create(string id)
        {
            log.Info("http get mehtod:" + id);
            SupplierManage supplierService = new SupplierManage();
            SupplierDetail singleForm = new SupplierDetail();
            supplierService.getSupplierBySupId(id);
            singleForm.sup = supplierService.supplier;
            singleForm.contactItem = supplierService.contactList;
            log.Debug("Supplier ID:" + singleForm.sup.SUPPLIER_ID);
            return View(singleForm);
        }

        public ActionResult AddSupplier(TND_SUPPLIER sup)
        {
            log.Info("create supplier process! supplier =" + sup.ToString());
            SupplierManage supplierService = new SupplierManage();
            SYS_USER u = (SYS_USER)Session["user"];
            string message = "";
            //1.新增供應商
            if (sup.SUPPLIER_ID == "" || sup.SUPPLIER_ID == null)
            {
                //新增供應商編號
                supplierService.newSupplier(sup);
                message = "輸入供應商基本資料 ! 若有聯絡人請先新增完聯絡人資料，再輸入供應商資料";
            }
            TempData["result"] = message;
            return Redirect("Create?id=" + sup.SUPPLIER_ID);
        }

        /// <summary>
        /// 取得聯絡人詳細資料
        /// </summary>
        /// <param name="contactid"></param>
        /// <returns></returns>
        public string getContact(string supplierid)
        {
            SupplierManage supplierService = new SupplierManage();
            log.Info("get contact info by supplier id=" + supplierid);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(supplierService.getContactById(supplierid));
            log.Info("supplier's Info=" + itemJson);
            return itemJson;
        }
        //新增供應聯絡人
        public String addContact(FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "新增聯絡人成功!!";

            TND_SUP_CONTACT_INFO item = new TND_SUP_CONTACT_INFO();
            item.SUPPLIER_MATERIAL_ID = form["supplier_id"];
            item.CONTACT_NAME = form["contact_name"];
            item.CONTACT_TEL = form["contact_tel"];
            item.CONTACT_FAX = form["contact_fax"];
            item.CONTACT_MOBIL = form["contact_mobil"];
            item.CONTACT_EMAIL = form["contact_email"];
            item.REMARK = form["remark"];
            SupplierManage supplierService = new SupplierManage();
            int i = supplierService.refreshContact(item);
            if (i == 0) { msg = supplierService.message; }
            return msg;
        }

        //更新供應商資料
        public String RefreshSupplier(string id, FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "";
            // 取得供應商資料
            TND_SUPPLIER sup = new TND_SUPPLIER();
            sup.SUPPLIER_ID = form.Get("supplierid").Trim();
            sup.COMPANY_NAME = form.Get("company_name").Trim();
            sup.COMPANY_ID = form.Get("company_id").Trim();
            sup.CONTACT_ADDRESS = form.Get("contact_address").Trim();
            sup.REGISTER_ADDRESS = form.Get("register_address").Trim();
            sup.TYPE_MAIN = form.Get("type_main").Trim();
            try
            {
                sup.TYPE_SUB = int.Parse(form.Get("type_sub").Trim());
            }
            catch (Exception ex)
            {
                log.Error(ex.StackTrace);
            }
            sup.SUPPLY_NOTE = form.Get("supply_note").Trim();
            SupplierManage supplierService = new SupplierManage();
            string supplierid = form.Get("supplierid").Trim();
            if (form.Get("contactid") != null && "" != form.Get("contactid"))
            {
                string[] lstItemId = form.Get("contactid").Split(',');
                string[] lstName = form.Get("contactname").Split(',');
                string[] lstTel = form.Get("contacttel").Split(',');
                string[] lstFax = form.Get("contactfax").Split(',');
                string[] lstMobile = form.Get("contactmobil").Split(',');
                string[] lstEmail = form.Get("contactemail").Split(',');
                string[] lstRemark = form.Get("contactremark").Split(',');
                List<TND_SUP_CONTACT_INFO> lstItem = new List<TND_SUP_CONTACT_INFO>();
                for (int j = 0; j < lstItemId.Count(); j++)
                {
                    TND_SUP_CONTACT_INFO item = new TND_SUP_CONTACT_INFO();
                    item.CONTACT_ID = int.Parse(lstItemId[j]);
                    item.CONTACT_NAME = lstName[j];
                    item.CONTACT_TEL = lstTel[j];
                    item.CONTACT_FAX = lstFax[j];
                    item.CONTACT_MOBIL = lstMobile[j];
                    item.CONTACT_EMAIL = lstEmail[j];
                    item.REMARK = lstRemark[j];
                    lstItem.Add(item);
                }

                int i = supplierService.updateSupplier(supplierid, sup, lstItem);
                if (i == 0)
                {
                    msg = supplierService.message;
                }
                else
                {
                    msg = "更新/新增供應商資料成功，SUPPLIER_ID =" + supplierid;
                }

                log.Info("Request: SUPPLIER_ID = " + supplierid + "CONTACT_ID =" + form["contact_id"]);
                return msg;
            }
            int k = supplierService.updateOnlySupplier(supplierid, sup);
            if (k == 0)
            {
                msg = supplierService.message;
            }
            else
            {
                msg = "更新/新增供應商資料成功，SUPPLIER_ID =" + supplierid;
            }

            log.Info("Request: SUPPLIER_ID = " + supplierid + "CONTACT_ID =" + form["contact_id"]);
            return msg;
        }

    }
}

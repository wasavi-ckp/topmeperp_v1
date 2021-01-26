using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using log4net;
using Newtonsoft.Json;

using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Migrations;
using System.Data.SqlClient;
using System.IO;

using System.Text;
using topmeperp.Models;
using topmeperp.Service;

namespace topmeperp.Service
{
    public class CertifyService : ContextService
    {
        //ContextDeptService.cs
        //static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        //public ENT_DEPARTMENT department = null;
        //public List<ENT_DEPARTMENT> Org_department = null;

        //ContextService.cs
        public topmepEntities db;// = new topmepEntities();
        //定義上傳檔案存放路徑
        public static string strUploadPath = ConfigurationManager.AppSettings["UploadFolder"];
        public static string quotesFolder = "Quotes"; //廠商報價單路徑
        public static string projectMgrFolder = "Project"; //施工進度管理資料夾

        #region ContextService : ExecuteStoreQuery & getDataTable
        //Sample Code : It can get ADO.NET Dataset
        //public DataSet ExecuteStoreQuery(string sql, CommandType commandType, Dictionary<string, Object> parameters)
        //{
        //    var result = new DataSet();
        //     creates a data access context (DbContext descendant)
        //    using (var context = new topmepEntities())
        //    {
        //         creates a Command 
        //        var cmd = context.Database.Connection.CreateCommand();
        //        cmd.CommandType = commandType;
        //        cmd.CommandText = sql;
        //         adds all parameters
        //        foreach (var pr in parameters)
        //        {
        //            var p = cmd.CreateParameter();
        //            p.ParameterName = pr.Key;
        //            p.Value = pr.Value;
        //            cmd.Parameters.Add(p);
        //        }
        //        try
        //        {
        //             executes
        //            context.Database.Connection.Open();
        //            var reader = cmd.ExecuteReader();

        //             loop through all resultsets (considering that it's possible to have more than one)
        //            do
        //            {
        //                 loads the DataTable (schema will be fetch automatically)
        //                var tb = new DataTable();
        //                tb.Load(reader);
        //                result.Tables.Add(tb);

        //            } while (!reader.IsClosed);
        //        }
        //        finally
        //        {
        //             closes the connection
        //            context.Database.Connection.Close();
        //        }
        //    }
        //     returns the DataSet
        //    return result;
        //}
        //public DataTable getDataTable(string sql, Dictionary<string, Object> parameters)
        //{
        //    DataTable dt = null;
        //    using (var context = new topmepEntities())
        //    {
        //        SqlDataAdapter adapter = new SqlDataAdapter(sql, context.Database.Connection.ConnectionString);
        //        foreach (var pr in parameters)
        //        {
        //            SqlParameter p = new SqlParameter(pr.Key, pr.Value);
        //            adapter.SelectCommand.Parameters.Add(p);
        //        }
        //        DataSet ds = new DataSet();
        //        adapter.Fill(ds);
        //        dt = ds.Tables[0];
        //    }
        //    return dt;
        //}
        #endregion
        
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public PLAN_CERT_ORDER planCertOrder = null;
        public List<PLAN_SUP_INQUIRY> planSupInquiry = null;
        public List<PLAN_SUP_INQUIRY_ITEM> planSupInquiryItem = null;
        public List<PLAN_ITEM2_SUP_INQUIRY> planItem2SupInquiry = null;
        public TND_SUPPLIER selectSupplier = null;                                                          //2021/01/25 13:57 冠蒲增加
        public List<TND_SUPPLIER> supplierList = null;                                                      //2021/01/25 13:57 冠蒲增加
        string sno_key = "PROJ";
        public string strMessage = null;
        private string filename = @"c:\\LogFiles\\outputTable.txt";

        public List<PLAN_SUP_INQUIRY> getPlanSupInquiryByProjectId(string projectId)
        {
            try
            {
                using (var context = new topmepEntities())
                {
                    planSupInquiry = context.PLAN_SUP_INQUIRY.SqlQuery("SELECT * FROM PLAN_SUP_INQUIRY"
                                                                        //+ " WHERE EXIST (SELECT * FROM PLAN_ITEM2_SUP_INQUIRY WHERE PLAN_ITEM2_SUP_INQUIRY.INQUIRY_FORM_ID LIKE PLAN_SUP_INQUIRY.INQUIRY_FORM_ID)"
                                                                        + " WHERE (SUPPLIER_ID IS NOT NULL) AND (PROJECT_ID = @projectId)"
                                                                        , new SqlParameter("projectId", projectId)).ToList();
                }
                //using (var context = new topmepEntities())
                //{
                //    planItem2SupInquiry = context.PLAN_ITEM2_SUP_INQUIRY.SqlQuery("SELECT * FROM PLAN_ITEM2_SUP_INQUIRY"
                //                                                                + " WHERE PROJECT_ID = @projectId"
                //                                                                , new SqlParameter("projectId", projectId)).ToList();
                //}
                //planSupInquiry = planSupInquiry.Where("PLAN_SUP_INQUIRY.INQUIRY_FORM_ID= PLAN_ITEM2_SUP_INQUIRY.INQUIRY_FORM_ID");
            }
            catch (Exception ex)
            {
                logger.Error(projectId + " CertifyService.getPlanSupInquiryByProjectId : " + ex.Message);
            }
            return planSupInquiry;
            
        }

        public List<PLAN_SUP_INQUIRY_ITEM> getPlanSupInquiryItemByInquiryFormId(string Iqfid)
        {
            try
            {
                using (var context = new topmepEntities())
                {
                    planSupInquiryItem = context.PLAN_SUP_INQUIRY_ITEM.SqlQuery("SELECT * FROM PLAN_SUP_INQUIRY_ITEM"
                                                                        //+ " RIGHT JOIN PLAN_ITEM2_SUP_INQUIRY ON PLAN_ITEM2_SUP_INQUIRY.INQUIRY_FORM_ID = PLAN_SUP_INQUIRY.INQUIRY_FORM_ID"
                                                                        //+ " WHERE (SUPPLIER_ID IS NOT NULL)"
                                                                        //+ " AND (STATUS <> '註銷')"
                                                                        + " WHERE INQUIRY_FORM_ID = @Iqfid"
                                                                        , new SqlParameter("Iqfid", Iqfid)).ToList();
                }
            }
            catch (Exception ex)
            {
                logger.Error(Iqfid + " CertifyService.getPlanSupInquiryItemByInquiryFormId : " + ex.Message);
            }
            return planSupInquiryItem;
        }
    }
}
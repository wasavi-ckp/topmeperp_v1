using log4net;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    //成本預算管制表Service Layer
    public class ContextService4PlanCost : PlanService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public CostControlInfo CostInfo = new CostControlInfo();
        public ContextService4PlanCost()
        {

        }
        public void getCostControlInfo(string projectId)
        {
            logger.Debug("get Cost Controll Info By ProjectId=" + projectId);
            //專案基本資料
            CostInfo.Project = getProject(projectId);
            //1.合約金額與追加減項目
            CostInfo.Revenue = getPlanRevenueById(projectId);
            //1.1 異動單彙整資料
            CostInfo.lstCostChangeEvent = getCostChangeEvnet(projectId);
            //2.直接成本:材料與工資
            PurchaseFormService pfservice = new PurchaseFormService();
            CostInfo.lstDirectCostItem = pfservice.getPlanContract(projectId);
            CostInfo.lstDirectCostItem.AddRange(pfservice.getPlanContract4Wage(projectId));
            //3.間接成本
            CostInfo.lstIndirectCostItem = getIndirectCost();
        }
        //建立間接成本資料
        public void createIndirectCost(string projectId, string userid)
        {
            List<SYS_PARA> lstItem = SystemParameter.getSystemPara("IndirectCostItem");
            List<PLAN_INDIRECT_COST> lstIndirectCostItem = new List<PLAN_INDIRECT_COST>();
            //取得合約金額
            CostInfo.Revenue = getPlanRevenueById(projectId);
            //取得間接成本項目
            foreach (SYS_PARA p in lstItem)
            {
                PLAN_INDIRECT_COST it = new PLAN_INDIRECT_COST();
                it.PROJECT_ID = projectId;
                it.FIELD_ID = p.FIELD_ID;
                it.FIELD_DESC = p.VALUE_FIELD;
                it.PERCENTAGE = decimal.Parse(p.KEY_FIELD);
                it.MODIFY_ID = userid;
                it.MODIFY_DATE = DateTime.Now;
                // System.Convert.ToDoublSystem.Math.Round(1.235, 2, MidpointRounding.AwayFromZero)
                it.COST = Convert.ToDecimal(Math.Round(Convert.ToDouble(CostInfo.Revenue.PLAN_REVENUE * decimal.Parse(p.KEY_FIELD) / 100), 0, MidpointRounding.AwayFromZero));
                logger.Debug(p.VALUE_FIELD + " Indirect Cost=" + it.COST + ",per=" + p.KEY_FIELD);
                lstIndirectCostItem.Add(it);
            }
            using (var context = new topmepEntities())
            {
                ///刪除現有資料
                string sql = "DELETE FROM PLAN_INDIRECT_COST WHERE PROJECT_ID=@projectId";
                logger.Debug("sql=" + sql + ",projectid=" + projectId);
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectId", projectId));
                context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                logger.Debug("sql=" + sql + ",projectid=" + projectId);
                ///將新資料存入
                context.PLAN_INDIRECT_COST.AddRange(lstIndirectCostItem);
                context.SaveChanges();
            }
        }
        //取得間接成本資料
        private List<PLAN_INDIRECT_COST> getIndirectCost()
        {
            List<PLAN_INDIRECT_COST> lst = null;
            using (var context = new topmepEntities())
            {
                string sql = "SELECT * FROM PLAN_INDIRECT_COST WHERE PROJECT_ID=@projectId ORDER BY FIELD_ID";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectId", CostInfo.Project.PROJECT_ID));
                logger.Debug("SQL:" + sql + ",projectId=" + CostInfo.Project.PROJECT_ID);
                lst = context.PLAN_INDIRECT_COST.SqlQuery(sql, parameters.ToArray()).ToList();
            }
            return lst;
        }
        //修正間接成本資料
        public void modifyIndirectCost(string projectId, List<PLAN_INDIRECT_COST> items)
        {
            using (var context = new topmepEntities())
            {
                ///逐筆更新資料
                string sql = "UPDATE PLAN_INDIRECT_COST SET COST = @cost, MODIFY_ID = @modifyId, MODIFY_DATE = @modifyDate, NOTE = ISNULL(Note,'') + @Note  WHERE PROJECT_ID = @projectId AND FIELD_ID = @fieldId";
                logger.Debug("sql=" + sql + ",projectid=" + projectId);
                foreach (PLAN_INDIRECT_COST it in items)
                {
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("projectId", projectId));
                    parameters.Add(new SqlParameter("fieldId", it.FIELD_ID));
                    parameters.Add(new SqlParameter("cost", it.COST));
                    parameters.Add(new SqlParameter("modifyId", it.MODIFY_ID));
                    parameters.Add(new SqlParameter("modifyDate", DateTime.Now));
                    parameters.Add(new SqlParameter("Note", it.NOTE));
                    context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                    logger.Debug("sql=" + sql + ",projectid=" + projectId);
                }
                ///將新資料存入
                context.SaveChanges();
            }
        }

        ///成本異動彙整資料(對業主)
        public List<CostChangeEvent> getCostChangeEvnet(string projectId)
        {
            List<CostChangeEvent> lstForms = new List<CostChangeEvent>();
            //2.取得異動單彙整資料
            using (var context = new topmepEntities())
            {
                //僅針對追加減部分列入 TRANSFLAG='1'
                logger.Debug("query by project and remark:" + projectId);
                string sql = @"SELECT FORM_ID,(REMARK_ITEM + REMARK_QTY + REMARK_PRICE + REMARK_OTHER) REMARK,SETTLEMENT_DATE,
                            (SELECT SUM(ITEM_QUANTITY * ITEM_UNIT_PRICE) FROM PLAN_COSTCHANGE_ITEM WHERE FORM_ID = f.FORM_ID ) TotalAmt,
                            (SELECT SUM(ITEM_QUANTITY * ITEM_UNIT_PRICE) FROM PLAN_COSTCHANGE_ITEM WHERE FORM_ID = f.FORM_ID AND TRANSFLAG='1') RecognizeAmt,
                            (SELECT SUM(ITEM_QUANTITY * ITEM_UNIT_PRICE) FROM[PLAN_COSTCHANGE_ITEM] WHERE FORM_ID = f.FORM_ID AND ITEM_QUANTITY> 0 AND TRANSFLAG='1') AddAmt,
                            (SELECT SUM(ITEM_QUANTITY * ITEM_UNIT_PRICE) FROM[PLAN_COSTCHANGE_ITEM] WHERE FORM_ID = f.FORM_ID AND ITEM_QUANTITY< 0 AND TRANSFLAG='1') CutAmt
                            FROM PLAN_COSTCHANGE_FORM f WHERE PROJECT_ID=@projectId AND STATUS IN ('01','02'); ";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectId", projectId));
                logger.Debug("SQL:" + sql);
                lstForms = context.Database.SqlQuery<CostChangeEvent>(sql, parameters.ToArray()).ToList();
            }
            return lstForms;
        }
    }
    //成本異動Service Layer
    public class CostChangeService : PlanService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string strSerialNoKey = "CC";
        public PLAN_COSTCHANGE_FORM form = null;
        public List<PLAN_COSTCHANGE_ITEM> lstItem = null;
        //建立異動單
        public string createChangeOrder(PLAN_COSTCHANGE_FORM form, List<PLAN_COSTCHANGE_ITEM> lstItem)
        {
            //1.新增成本異動單
            SerialKeyService skService = new SerialKeyService();
            form.FORM_ID = skService.getSerialKey(strSerialNoKey);
            form.STATUS = "10";
            PLAN_ITEM pi = null;
            int i = 0;
            //2.將資料寫入 
            using (var context = new topmepEntities())
            {
                context.PLAN_COSTCHANGE_FORM.Add(form);
                logger.Debug("create COSTCHANGE_FORM:" + form.FORM_ID);
                foreach (PLAN_COSTCHANGE_ITEM item in lstItem)
                {
                    if (null != item.PLAN_ITEM_ID && "" != item.PLAN_ITEM_ID)
                    {
                        logger.Debug("Object in contract :" + item.PLAN_ITEM_ID);
                        pi = context.PLAN_ITEM.SqlQuery("SELECT * FROM PLAN_ITEM WHERE PLAN_ITEM_ID=@itemId", new SqlParameter("itemId", item.PLAN_ITEM_ID)).First();
                        //補足標單品項欄位
                        if (pi != null && item.ITEM_ID == null)
                        {
                            item.ITEM_ID = pi.ITEM_ID;
                        }
                        if (pi != null && item.ITEM_DESC == null)
                        {
                            item.ITEM_DESC = pi.ITEM_DESC;
                        }
                        if (pi != null && item.ITEM_UNIT == null)
                        {
                            item.ITEM_UNIT = pi.ITEM_UNIT;
                        }
                        if (pi != null && item.ITEM_UNIT_PRICE == null)
                        {
                            item.ITEM_UNIT_PRICE = pi.ITEM_UNIT_PRICE;
                        }
                        if (pi != null && item.ITEM_UNIT_COST == null)
                        {
                            item.ITEM_UNIT_COST = pi.ITEM_UNIT_COST;
                        }
                    }
                    item.FORM_ID = form.FORM_ID;
                    item.PROJECT_ID = form.PROJECT_ID;
                    context.PLAN_COSTCHANGE_ITEM.Add(item);
                    item.CREATE_USER_ID = form.CREATE_USER_ID;
                    item.CREATE_DATE = form.CREATE_DATE;
                    logger.Debug("create COSTCHANGE_ITEM:" + item.PLAN_ITEM_ID);
                }
                i = context.SaveChanges();
            }
            logger.Info("add CostChangeItem count =" + i);
            return form.FORM_ID;
        }

        //取得單一異動單資料
        public PLAN_COSTCHANGE_FORM getChangeOrderForm(string formId)
        {
            //2.取得異動單資料
            using (var context = new topmepEntities())
            {
                logger.Debug("change form Id:" + formId);
                string sql = "SELECT * FROM PLAN_COSTCHANGE_FORM WHERE FORM_ID=@formId";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("formId", formId));
                logger.Debug("SQL:" + sql);
                form = context.PLAN_COSTCHANGE_FORM.SqlQuery(sql, parameters.ToArray()).First();
                lstItem = form.PLAN_COSTCHANGE_ITEM.ToList();
            }
            project = getProject(form.PROJECT_ID);
            return form;
        }
        //更新異動單資料
        public string updateChangeOrder(PLAN_COSTCHANGE_FORM form, List<PLAN_COSTCHANGE_ITEM> lstItem)
        {
            int i = 0;
            string sqlForm = @"UPDATE PLAN_COSTCHANGE_FORM SET REASON_CODE=@Reasoncode,METHOD_CODE=@methodCode,
                            REMARK_ITEM=@RemarkItem,REMARK_QTY=Null,REMARK_PRICE=Null,REMARK_OTHER=Null,
                            MODIFY_USER_ID=@userId,MODIFY_DATE=@modifyDate WHERE FORM_ID=@formId;";
            string sqlItem = @"UPDATE PLAN_COSTCHANGE_ITEM SET ITEM_DESC=@itemdesc,ITEM_UNIT=@unit,ITEM_UNIT_PRICE=@unitPrice,ITEM_UNIT_COST=@unitCost,
                              ITEM_QUANTITY=@Qty,ITEM_REMARK=@remark,TRANSFLAG=@transFlag,MODIFY_USER_ID=@userId,MODIFY_DATE=@modifyDate 
                              WHERE ITEM_UID=@uid";
            //2.將資料寫入 

            using (var context = new topmepEntities())
            {
                try
                {
                    //更新表頭
                    context.Database.BeginTransaction();
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("Reasoncode", form.REASON_CODE));
                    parameters.Add(new SqlParameter("methodCode", form.METHOD_CODE));
                    parameters.Add(new SqlParameter("RemarkItem", form.REMARK_ITEM));
                   // parameters.Add(new SqlParameter("RemarkQty", form.REMARK_QTY));
                   // parameters.Add(new SqlParameter("RemarkPrice", form.REMARK_PRICE));
                   // parameters.Add(new SqlParameter("RemarkOther", form.REMARK_OTHER));
                    parameters.Add(new SqlParameter("userId", form.MODIFY_USER_ID));
                    parameters.Add(new SqlParameter("modifyDate", form.MODIFY_DATE));
                    parameters.Add(new SqlParameter("formId", form.FORM_ID));
                    i = context.Database.ExecuteSqlCommand(sqlForm, parameters.ToArray());
                    logger.Debug("create COSTCHANGE_FORM:" + sqlForm);
                    foreach (PLAN_COSTCHANGE_ITEM item in lstItem)
                    {
                        parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("itemdesc", item.ITEM_DESC));
                        parameters.Add(new SqlParameter("unit", item.ITEM_UNIT));
                        if (item.ITEM_UNIT_PRICE != null)
                        {
                            parameters.Add(new SqlParameter("unitPrice", item.ITEM_UNIT_PRICE));
                        }
                        else
                        {
                            parameters.Add(new SqlParameter("unitPrice", DBNull.Value));
                        }
                        if (item.ITEM_UNIT_COST != null)
                        {
                            parameters.Add(new SqlParameter("unitCost", item.ITEM_UNIT_COST));
                        }
                        else
                        {
                            parameters.Add(new SqlParameter("unitCost", DBNull.Value));
                        }
                        if (item.ITEM_QUANTITY == null)
                        {
                            parameters.Add(new SqlParameter("Qty", DBNull.Value));
                        }
                        else
                        {
                            parameters.Add(new SqlParameter("Qty", item.ITEM_QUANTITY));
                        }
                        parameters.Add(new SqlParameter("transFlag", item.TRANSFLAG));
                        parameters.Add(new SqlParameter("remark", item.ITEM_REMARK));
                        parameters.Add(new SqlParameter("userId", form.MODIFY_USER_ID));
                        parameters.Add(new SqlParameter("modifyDate", form.MODIFY_DATE));
                        parameters.Add(new SqlParameter("uid", item.ITEM_UID));
                        i = i + context.Database.ExecuteSqlCommand(sqlItem, parameters.ToArray());
                    }
                    context.Database.CurrentTransaction.Commit();
                }
                catch (Exception ex)
                {
                    context.Database.CurrentTransaction.Rollback();
                    logger.Error(ex.Message + ":" + ex.StackTrace);
                    return "資料更新失敗!!(" + ex.Message + ")";
                }
            }

            return "資料更新成功(" + i + ")!";
        }
        //新增異動單品項
        public int addChangeOrderItem(PLAN_COSTCHANGE_ITEM item)
        {
            int i = 0;
            //2.將資料寫入 
            using (var context = new topmepEntities())
            {
                try
                {
                    logger.Debug("create COSTCHANGE_FORM:" + item.FORM_ID);
                    context.PLAN_COSTCHANGE_ITEM.Add(item);
                    i = context.SaveChanges();
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":" + ex.StackTrace);
                }
            }
            return i;
        }
        //移除異動單品項
        public int delChangeOrderItem(long itemid)
        {
            int i = 0;
            //2.將品項資料刪除
            using (var context = new topmepEntities())
            {
                try
                {
                    string sql = "DELETE FROM PLAN_COSTCHANGE_ITEM WHERE ITEM_UID=@itemUid;";
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("itemUid", itemid));
                    logger.Debug("Delete COSTCHANGE_ITEM:" + itemid);
                    context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                    i = context.SaveChanges();
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":" + ex.StackTrace);
                }
            }
            return i;
        }
        //取得成本異動單資料
        public List<CostChangeForm> getCostChangeForm(string projectId, string status,string remark,string noInquiry)
        {
            logger.Debug("get costchange form for project id=" + projectId);
            List<CostChangeForm> lst = null;
            string sql = @"SELECT *,
                    (SELECT  VALUE_FIELD FROM SYS_PARA
                    WHERE FUNCTION_ID='COSTHANGE' AND FIELD_ID='REASON' AND KEY_FIELD=F.REASON_CODE ) REASON,
                    (SELECT  VALUE_FIELD FROM SYS_PARA　
                    WHERE FUNCTION_ID='COSTHANGE' AND FIELD_ID='METHOD' AND KEY_FIELD=F.METHOD_CODE ) METHOD
                     FROM PLAN_COSTCHANGE_FORM F WHERE PROJECT_ID=@projectId ";
            if (null == noInquiry)
            {
                sql = sql + " AND INQUIRY_FORM_ID IS NULL";
            } else
            {
                
            }
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectId", projectId));
            if (null != status && status != "")
            {
                sql = sql + " AND STATUS=@status";
                parameters.Add(new SqlParameter("status", status));
            }
            if (null != remark && remark != "")
            {
                sql = sql + " AND (REMARK_ITEM Like @remark OR REMARK_QTY Like @remark OR REMARK_PRICE Like @remark  OR REMARK_OTHER Like @remark) ";
                parameters.Add(new SqlParameter("remark", "%" + remark + "%"));
            }
            logger.Debug("sql" + sql);
            using (var context = new topmepEntities())
            {
                try
                {
                    lst = context.Database.SqlQuery<CostChangeForm>(sql, parameters.ToArray()).ToList();
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + "," + ex.StackTrace);
                }
            }
            return lst;
        }
        /// <summary>
        /// 依據成本異動單建立詢價單
        /// </summary>
        public string createInquiryOrderByChangeForm(string formId,SYS_USER u)
        {
            //取得異動單資料
            getChangeOrderForm(formId);
            logger.Info("create new [PLAN_SUP_INQUIRY] from CostChange Order= "+ formId);
            string sno_key = "PC";
            SerialKeyService snoservice = new SerialKeyService();

            using (var context = new topmepEntities())
            {
                //1.取得異動單相關資訊
                //2,建立表頭
                PLAN_SUP_INQUIRY pf = new PLAN_SUP_INQUIRY();
                pf.INQUIRY_FORM_ID = snoservice.getSerialKey(sno_key);
                pf.PROJECT_ID = form.PROJECT_ID;
                pf.FORM_NAME = "(成本異動)" + form.REMARK_ITEM;
                //聯絡人基本資料
                pf.OWNER_NAME = u.USER_NAME;
                pf.OWNER_EMAIL = u.EMAIL;
                pf.OWNER_TEL = u.TEL + "-" + u.TEL_EXT;
                pf.OWNER_FAX = u.FAX;
                pf.CREATE_ID = u.USER_ID;
                pf.CREATE_DATE = DateTime.Now;

                context.PLAN_SUP_INQUIRY.Add(pf);
                int i = context.SaveChanges();
                logger.Info("plan form id = " + pf.INQUIRY_FORM_ID);               
                //3建立詢價單明細
                string sql = @"INSERT INTO PLAN_SUP_INQUIRY_ITEM 
                        (INQUIRY_FORM_ID,PLAN_ITEM_ID,ITEM_ID,ITEM_DESC,ITEM_UNIT,ITEM_QTY,ITEM_UNIT_PRICE,ITEM_REMARK)
                        SELECT @InquiryFormId INQUIRY_FORM_ID,PLAN_ITEM_ID,ITEM_ID,ITEM_DESC,ITEM_UNIT,ITEM_QUANTITY ITEM_QTY,ITEM_UNIT_COST ITEM_UNICE_PRICE,ITEM_REMARK 
                         FROM PLAN_COSTCHANGE_ITEM WHERE FORM_ID=@formId";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("InquiryFormId", pf.INQUIRY_FORM_ID));
                parameters.Add(new SqlParameter("formId", formId));
                i = context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                //4.將成本異動單更新相關的詢價單資料
                sql = @"UPDATE PLAN_COSTCHANGE_FORM SET INQUIRY_FORM_ID=@InquiryFormId,
                            MODIFY_USER_ID = @userId, MODIFY_DATE = @modifyDate WHERE FORM_ID = @formId; ";
                logger.Info("sql =" + sql);
                parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("InquiryFormId", pf.INQUIRY_FORM_ID));
                parameters.Add(new SqlParameter("formId", formId));
                parameters.Add(new SqlParameter("userId", u.USER_ID));
                parameters.Add(new SqlParameter("modifyDate",DateTime.Now));
                i = context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                return pf.INQUIRY_FORM_ID;
            }
        }
    }
}

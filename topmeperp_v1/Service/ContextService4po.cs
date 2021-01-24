using log4net;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Core.Objects;
using System.Data.Entity.Migrations;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using topmeperp.Models;
using System.Globalization;
using System.Text;

namespace topmeperp.Service
{
    public class PlanService : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public string message = "";
        public TND_PROJECT project = null;
        public TND_PROJECT budgetTable = null;
        /// <summary>
        /// 取得工地專案資料
        /// </summary>
        /// <param name="projectname"></param>
        /// <param name="status"></param>
        /// <returns></returns>
        public static List<ProjectList> SearchProjectByName(string projectname, string status, SYS_USER u)
        {
            if (projectname != null)
            {
                logger.Info("search project by 名稱 =" + projectname);
                List<ProjectList> lstProject = new List<ProjectList>();
                string sql = @"select DISTINCT p.*, convert(varchar, pi.CREATE_DATE , 111) as PLAN_CREATE_DATE 
                        from TND_PROJECT p left join 
                        (SELECT PROJECT_ID, MIN(CREATE_DATE) AS CREATE_DATE 
                        FROM PLAN_ITEM GROUP BY PROJECT_ID)pi on p.PROJECT_ID = pi.PROJECT_ID 
                        where p.PROJECT_NAME Like '%' + @projectname + '%' 
                        AND STATUS  IN ('" + @status + "') AND $UserCond p.PROJECT_ID !='001' ORDER BY STATUS DESC";
                var parameters = new List<SqlParameter>();

                parameters.Add(new SqlParameter("projectname", projectname));
                parameters.Add(new SqlParameter("status", status));
                if (null != u)
                {
                    //加入使用者專案權限
                    string userCond = " p.PROJECT_ID IN (SELECT PROJECT_ID FROM TND_TASKASSIGN WHERE USER_ID=@UserID) AND ";
                    sql = sql.Replace("$UserCond", userCond);
                    parameters.Add(new SqlParameter("UserID", u.USER_NAME));
                }
                else
                {
                    sql = sql.Replace("$UserCond", "");
                }
                using (var context = new topmepEntities())
                {
                    lstProject = context.Database.SqlQuery<ProjectList>(sql, parameters.ToArray()).ToList();
                }
                logger.Info("get project count=" + lstProject.Count);
                return lstProject;
            }
            else
            {
                return null;
            }
        }
        #region 得標標單項目處理
        public TND_PROJECT getProject(string prjid)
        {
            using (var context = new topmepEntities())
            {
                project = context.TND_PROJECT.SqlQuery("select p.* from TND_PROJECT p "
                    + "where p.PROJECT_ID = @pid "
                   , new SqlParameter("pid", prjid)).FirstOrDefault();
            }
            return project;
        }
        #endregion

        public string getBudgetById(string prjid)
        {
            string projectid = null;
            using (var context = new topmepEntities())
            {
                projectid = context.Database.SqlQuery<string>("select DISTINCT PROJECT_ID FROM PLAN_BUDGET WHERE PROJECT_ID = @pid "
               , new SqlParameter("pid", prjid)).FirstOrDefault();
            }
            return projectid;
        }

        budgetsummary budget = null;
        public budgetsummary getBudgetForComparison(string projectid, string typecode1, string typecode2, string systemMain, string systemSub)
        {
            using (var context = new topmepEntities())
            {
                if (null != typecode1 && typecode1 != "" && typecode2 == "" || null != typecode1 && typecode1 != "" && typecode2 == null)
                {
                    budget = context.Database.SqlQuery<budgetsummary>("SELECT TYPE_CODE_1,SUM(BUDGET_AMOUNT) AS BAmount " +
                        "FROM PLAN_BUDGET WHERE PROJECT_ID = @pid GROUP BY TYPE_CODE_1 HAVING TYPE_CODE_1 = @typecode1 "
                       , new SqlParameter("pid", projectid), new SqlParameter("typecode1", typecode1)).First();
                }
                else if (null != typecode1 && typecode1 != "" && null != typecode2 && typecode2 != "")
                {
                    budget = context.Database.SqlQuery<budgetsummary>("SELECT TYPE_CODE_1, TYPE_CODE_1, BUDGET_AMOUNT AS BAmount " +
                        "FROM PLAN_BUDGET WHERE PROJECT_ID = @pid AND TYPE_CODE_1 = @typecode1 AND TYPE_CODE_2 = @typecode2 "
                       , new SqlParameter("pid", projectid), new SqlParameter("typecode1", typecode1), new SqlParameter("typecode2", typecode2)).FirstOrDefault();
                }
                else if (null == typecode1 && null == typecode2 && null == systemMain && null == systemSub || typecode1 == "" && typecode2 == "" && systemMain == "" && systemSub == "")
                {
                    budget = context.Database.SqlQuery<budgetsummary>("SELECT SUM(BUDGET_AMOUNT) AS BAmount " +
                        "FROM PLAN_BUDGET WHERE PROJECT_ID = @pid "
                       , new SqlParameter("pid", projectid)).First();
                }
                else
                {
                    budget = null;
                }
            }
            return budget;
        }
        public int addBudget(List<PLAN_BUDGET> lstItem)
        {
            //1.新增預算資料
            int i = 0;
            logger.Info("add budget = " + lstItem.Count);
            //2.將預算資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_BUDGET item in lstItem)
                {
                    item.CREATE_DATE = DateTime.Now;
                    context.PLAN_BUDGET.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add budget count =" + i);
            return i;
        }

        public int updateBudget(string projectid, List<PLAN_BUDGET> lstItem)
        {
            //1.修改預算資料
            int i = 0;
            logger.Info("update budget = " + lstItem.Count);
            //2.將預算資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_BUDGET item in lstItem)
                {
                    PLAN_BUDGET existItem = null;
                    logger.Debug("plan budget id=" + item.PLAN_BUDGET_ID);
                    if (item.PLAN_BUDGET_ID != 0)
                    {
                        existItem = context.PLAN_BUDGET.Find(item.PLAN_BUDGET_ID);
                    }
                    else
                    {
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("projectid", projectid));
                        parameters.Add(new SqlParameter("code1", item.TYPE_CODE_1));
                        parameters.Add(new SqlParameter("code2", item.TYPE_CODE_2));

                        string sql = "SELECT * FROM PLAN_BUDGET WHERE PROJECT_ID = @projectid and ISNULL(TYPE_CODE_1, '') + ISNULL(TYPE_CODE_2, '') = @code1 + @code2";
                        logger.Info(sql + " ;" + item.PROJECT_ID + item.TYPE_CODE_1 + item.TYPE_CODE_2 + item.SYSTEM_MAIN + item.SYSTEM_SUB);
                        PLAN_BUDGET excelItem = context.PLAN_BUDGET.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_BUDGET.Find(excelItem.PLAN_BUDGET_ID);

                    }
                    logger.Debug("find exist item=" + existItem.PLAN_BUDGET_ID);
                    existItem.BUDGET_RATIO = item.BUDGET_RATIO;
                    existItem.BUDGET_WAGE_RATIO = item.BUDGET_WAGE_RATIO;
                    existItem.MODIFY_ID = item.MODIFY_ID;
                    existItem.MODIFY_DATE = DateTime.Now;
                    context.PLAN_BUDGET.AddOrUpdate(existItem);
                }
                i = context.SaveChanges();
            }
            logger.Info("update budget count =" + i);
            return i;
        }
        public int updateBudgetToPlanItem(string id)
        {
            int i = 0;
            logger.Info("update budget ratio to plan items by id :" + id);
            string sql = "UPDATE PLAN_ITEM SET PLAN_ITEM.BUDGET_RATIO = plan_budget.BUDGET_RATIO, PLAN_ITEM.BUDGET_WAGE_RATIO = plan_budget.BUDGET_WAGE_RATIO " +
                   "from PLAN_ITEM inner join " +
                   "plan_budget on REPLACE(PLAN_ITEM.PROJECT_ID, ' ', '') + IIF(REPLACE(PLAN_ITEM.TYPE_CODE_1, ' ', '') is null, '', IIF(REPLACE(PLAN_ITEM.TYPE_CODE_1, ' ', '') = 0, '', REPLACE(PLAN_ITEM.TYPE_CODE_1, ' ', ''))) + IIF(REPLACE(PLAN_ITEM.TYPE_CODE_2, ' ', '') is null, '', IIF(REPLACE(PLAN_ITEM.TYPE_CODE_2, ' ', '') = 0, '', REPLACE(PLAN_ITEM.TYPE_CODE_2, ' ', ''))) " +
                   "= REPLACE(plan_budget.PROJECT_ID, ' ', '') + IIF(REPLACE(plan_budget.TYPE_CODE_1, ' ', '') is null, '', IIF(REPLACE(plan_budget.TYPE_CODE_1, ' ', '') = 0, '', REPLACE(plan_budget.TYPE_CODE_1, ' ', ''))) + IIF(REPLACE(plan_budget.TYPE_CODE_2, ' ', '') is null, '', IIF(REPLACE(plan_budget.TYPE_CODE_2, ' ', '') = 0, '', REPLACE(plan_budget.TYPE_CODE_2, ' ', ''))) WHERE PLAN_ITEM.PROJECT_ID  = @id ";
            logger.Debug("sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("id", id));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return i;
        }
        //更新得標標單品項個別預算
        public int updateItemBudget(string projectid, List<PLAN_ITEM> lstItem)
        {
            //1.新增預算資料
            int i = 0;
            logger.Info("update budget = " + lstItem.Count);
            //2.將預算資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_ITEM item in lstItem)
                {
                    PLAN_ITEM existItem = null;
                    logger.Debug("plan item id=" + item.PLAN_ITEM_ID);
                    if (item.PLAN_ITEM_ID != null && item.PLAN_ITEM_ID != "")
                    {
                        existItem = context.PLAN_ITEM.Find(item.PLAN_ITEM_ID);
                    }
                    else
                    {
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("projectid", projectid));
                        parameters.Add(new SqlParameter("planitemid", item.PLAN_ITEM_ID));
                        string sql = "SELECT * FROM PLAN_ITEM WHERE PROJECT_ID = @projectid and PLAN_ITEM_ID = @planitemid ";
                        logger.Info(sql + " ;" + item.PROJECT_ID + item.PLAN_ITEM_ID);
                        PLAN_ITEM excelItem = context.PLAN_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_ITEM.Find(excelItem.PLAN_ITEM_ID);
                    }
                    logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                    existItem.BUDGET_RATIO = item.BUDGET_RATIO;
                    existItem.MODIFY_USER_ID = item.MODIFY_USER_ID;
                    existItem.MODIFY_DATE = DateTime.Now;
                    context.PLAN_ITEM.AddOrUpdate(existItem);
                }
                i = context.SaveChanges();
            }
            logger.Info("update budget count =" + i);
            return i;
        }

        #region 取得預算表單檔頭資訊
        //取得預算表單檔頭
        public void getProjectId(string projectid)
        {
            logger.Info("get project : projectid=" + projectid);
            using (var context = new topmepEntities())
            {
                //取得預算表單檔頭資訊
                budgetTable = context.TND_PROJECT.SqlQuery("SELECT * FROM TND_PROJECT WHERE PROJECT_ID=@projectid", new SqlParameter("projectid", projectid)).First();
            }
        }
        #endregion

        #region 預算數量  
        //預算上傳數量  
        public int refreshBudget(List<PLAN_BUDGET> items)
        {
            //1.檢查專案是否存在
            //if (null == project) { throw new Exception("Project is not exist !!"); } 先註解掉,因為讀取不到project,會造成null == project is true,
            //而導致錯誤, 因為已設定是直接由專案頁面導入上傳圖算畫面，故不會有專案不存在的bug
            int i = 0;
            logger.Info("refreshBudgetItem = " + items.Count);
            //2.將Excel 資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_BUDGET item in items)
                {
                    //item.PROJECT_ID = project.PROJECT_ID;先註解掉,因為專案編號一開始已經設定了，會直接代入
                    context.PLAN_BUDGET.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add PLAN_BUDGET count =" + i);
            return i;
        }
        public int delBudgetByProject(string projectid)
        {
            logger.Info("remove all budget by project ID=" + projectid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete all PLAN_BUDGET by proejct id=" + projectid);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_BUDGET WHERE PROJECT_ID=@projectid", new SqlParameter("@projectid", projectid));
            }
            logger.Debug("delete PLAN_BUDGET count=" + i);
            return i;
        }
        #endregion

        PlanRevenue plan = null;
        public PlanRevenue getPlanRevenueById(string prjid)
        {
            using (var context = new topmepEntities())
            {
                string sql = @"SELECT p.PROJECT_ID AS CONTRACT_ID, 
                pcp.CONTRACT_PRODUCTION, 
                CONVERT(char(10), pcp.DELIVERY_DATE, 111) AS DELIVERY_DATE, 
                ISNULL(pcp.MAINTENANCE_BOND, 0) AS MAINTENANCE_BOND,
                CONVERT(char(10), pcp.MB_DUE_DATE, 111) AS MB_DUE_DATE, 
                pcp.REMARK AS ConRemark, 
                ppt.PAYMENT_ADVANCE_RATIO, 
                ppt.PAYMENT_RETENTION_RATIO, 
                (SELECT SUM(ITEM_UNIT_PRICE*ITEM_QUANTITY) FROM PLAN_ITEM pi WHERE pi.PROJECT_ID = @pid AND ISNULL(pi.IN_CONTRACT,'Y')='Y') AS PLAN_REVENUE 
                FROM TND_PROJECT p 
                LEFT JOIN PLAN_CONTRACT_PROCESS pcp ON p.PROJECT_ID = pcp.CONTRACT_ID 
                LEFT JOIN PLAN_PAYMENT_TERMS ppt ON p.PROJECT_ID = ppt.CONTRACT_ID WHERE p.PROJECT_ID = @pid ";
                logger.Debug("sql=" + sql + ",project_id=" + prjid);
                plan = context.Database.SqlQuery<PlanRevenue>(sql, new SqlParameter("pid", prjid)).First();
            }
            return plan;
        }

        public int delOwnerContractByProject(string projectId)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete owner contract by proejct id=" + projectId);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_CONTRACT_PROCESS WHERE PROJECT_ID=@projectid AND CONTRACT_ID =@projectid ", new SqlParameter("@projectid", projectId));
            }
            logger.Debug("delete owner contract count =" + i);
            return i;
        }

        //新增業主合約簽訂資料
        public int AddOwnerContractProcess(PLAN_CONTRACT_PROCESS contract)
        {
            string message = "";
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.PLAN_CONTRACT_PROCESS.AddOrUpdate(contract);
                    i = context.SaveChanges();
                }
                catch (Exception e)
                {
                    logger.Error("add owner contract process fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }
            }
            return i;
        }
        public int addContractId4Owner(string projectid)
        {
            int i = 0;
            //將業主合約編號寫入PLAN PAYMENT TERMS
            logger.Info("copy contract id from owner into plan payment terms, project id =" + projectid);
            using (var context = new topmepEntities())
            {
                PLAN_PAYMENT_TERMS lstItem = new PLAN_PAYMENT_TERMS();
                string sql = "INSERT INTO PLAN_PAYMENT_TERMS (CONTRACT_ID, PROJECT_ID, TYPE) " +
                       "SELECT '" + projectid + "' AS contractid, '" + projectid + "', 'O' FROM TND_PROJECT p WHERE p.PROJECT_ID = '" + projectid + "'  " +
                       "AND '" + projectid + "' NOT IN(SELECT ppt.CONTRACT_ID FROM PLAN_PAYMENT_TERMS ppt) ";
                logger.Info("sql =" + sql);
                i = context.Database.ExecuteSqlCommand(sql);
                return i;
            }
        }
        //取得上傳的業主合約檔案紀錄
        public List<RevenueFromOwner> getOwnerContractFileByPrjId(string projectid)
        {

            logger.Info(" get owner's contract file by project id =" + projectid);
            List<RevenueFromOwner> lstItem = new List<RevenueFromOwner>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT f.FILE_ID AS ITEM_UID, f.PROJECT_ID, f.FILE_UPLOAD_NAME, f.FILE_ACTURE_NAME, f.FILE_TYPE, " +
                "f.CREATE_DATE, ROW_NUMBER() OVER(ORDER BY FILE_ID) AS NO FROM TND_FILE f WHERE f.PROJECT_ID = @projectid AND " +
                "f.FILE_UPLOAD_NAME = '業主合約' ";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectid));
            using (var context = new topmepEntities())
            {
                logger.Debug("get owner's contract file sql=" + sql);
                lstItem = context.Database.SqlQuery<RevenueFromOwner>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get owner's contract file record count=" + lstItem.Count);
            return lstItem;
        }
    }
    //採發階段
    public class Bill4PurchService : TnderProjectService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public TND_PROJECT wageTable = null;
        public List<PROJECT_ITEM_WITH_WAGE> wageTableItem = null;
        public TND_PROJECT name = null;
        public Bill4PurchService()
        {
        }
        //取得工率表單
        public void getProjectId(string projectid)
        {
            logger.Info("get project : projectid=" + projectid);
            using (var context = new topmepEntities())
            {
                //取得專案檔頭資訊
                wageTable = context.TND_PROJECT.SqlQuery("SELECT * FROM TND_PROJECT WHERE PROJECT_ID=@projectid", new SqlParameter("projectid", projectid)).First();
                //取得合約標單明細
                string sql = @"SELECT DISTINCT i.PLAN_ITEM_ID PROJECT_ITEM_ID,i.PROJECT_ID,i.ITEM_ID ,i.ITEM_DESC,i.ITEM_UNIT,i.ITEM_QUANTITY,i.ITEM_UNIT_PRICE
                   ,i.MAN_PRICE ,i.ITEM_REMARK  ,i.TYPE_CODE_1 ,i.TYPE_CODE_2 ,i.SUB_TYPE_CODE  ,i.SYSTEM_MAIN ,i.SYSTEM_SUB
                   ,i.MODIFY_USER_ID  ,i.MODIFY_DATE ,i.CREATE_USER_ID   ,i.CREATE_DATE  ,i.EXCEL_ROW_ID  ,i.FORM_NAME  ,i.SUPPLIER_ID
                   ,i.BUDGET_RATIO  ,i.ITEM_FORM_QUANTITY   ,i.ITEM_UNIT_COST ,i.TND_RATIO
                   ,i.MAN_FORM_NAME    ,i.MAN_SUPPLIER_ID      ,i.LEAD_TIME      ,i.DEL_FLAG      ,i.INQUIRY_FORM_ID      ,i.MAN_FORM_ID
                   ,i.BUDGET_WAGE_RATIO,w.ratio,w.price,map.QTY as MAP_QTY,ISNULL(i.IN_CONTRACT,'Y') IN_CONTRACT FROM PLAN_ITEM i LEFT OUTER JOIN 
                   TND_WAGE w ON i.PLAN_ITEM_ID = w.PROJECT_ITEM_ID 
                    LEFT OUTER JOIN vw_MAP_MATERLIALIST map ON i.PLAN_ITEM_ID = map.PROJECT_ITEM_ID 
                   WHERE i.project_id = @projectid AND ISNULL(i.DEL_FLAG,'N')='N' ORDER BY i.EXCEL_ROW_ID; ";
                //加入其他條件供下載

                wageTableItem = context.Database.SqlQuery<PROJECT_ITEM_WITH_WAGE>(sql, new SqlParameter("projectid", projectid)).ToList();
                logger.Debug("get project item count:" + wageTableItem.Count);
            }
        }
    }

    public class BudgetDataService : CostAnalysisDataService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public List<DirectCost> getBudget(string projectid)
        {
            List<DirectCost> lstBudget = new List<DirectCost>();
            using (var context = new topmepEntities())
            {
                string sql = @"SELECT TYPE_CODE_1 MAINCODE,BUDGET_NAME MAINCODE_DESC,TYPE_CODE_2 SUB_CODE,'' as SUB_DESC ,
                        CONTRACT_AMOUNT AS CONTRACT_PRICE,
                        BUDGET_AMOUNT MATERIAL_COST_INMAP,  BUDGET_RATIO BUDGET, BUDGET_AMOUNT*BUDGET_RATIO/100 MATERIAL_BUDGET_INMAP,
                    	BUDGET_WAGE_AMOUNT MAN_DAY_INMAP,BUDGET_WAGE_RATIO BUDGET_WAGE,BUDGET_WAGE_AMOUNT *BUDGET_WAGE_RATIO/100 MAN_DAY_BUDGET_INMAP
                        FROM PLAN_BUDGET WHERE PROJECT_ID=@projectid";
                logger.Info("sql = " + sql + ",project_id=" + projectid);
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", projectid));
                lstBudget = context.Database.SqlQuery<DirectCost>(sql, parameters.ToArray()).ToList();
                logger.Info("Get Budget Info Record Count=" + lstBudget.Count);
            }
            return lstBudget;
        }
        //取得投標標單總直接成本
        public DirectCost getTotalCost(string projectid)
        {
            DirectCost lstTotalCost = null;
            using (var context = new topmepEntities())
            {
                string sql = "SELECT SUM(ISNULL(TND_COST,0)) AS TOTAL_COST, SUM(ISNULL(BUDGET,0)) AS MATERIAL_BUDGET, SUM(ISNULL(BUDGET_WAGE,0)) AS WAGE_BUDGET, SUM(ISNULL(BUDGET,0)) + SUM(ISNULL(BUDGET_WAGE,0)) AS TOTAL_BUDGET, SUM(ISNULL(P_COST,0)) AS TOTAL_P_COST FROM (SELECT(select TYPE_CODE_1 + TYPE_CODE_2 from REF_TYPE_MAIN WHERE  " +
                    "TYPE_CODE_1 + TYPE_CODE_2 = A.TYPE_CODE_1) MAINCODE, (select TYPE_DESC from REF_TYPE_MAIN WHERE  TYPE_CODE_1 + TYPE_CODE_2 = A.TYPE_CODE_1) MAINCODE_DESC, " +
                    "(select SUB_TYPE_ID from REF_TYPE_SUB WHERE  A.TYPE_CODE_1 + A.TYPE_CODE_2 = SUB_TYPE_ID) T_SUB_CODE, TYPE_CODE_2 SUB_CODE, " +
                    "(select TYPE_DESC from REF_TYPE_SUB WHERE  A.TYPE_CODE_1 + A.TYPE_CODE_2 = SUB_TYPE_ID) SUB_DESC, SUM(MapQty * tndPrice) MATERIAL_COST_INMAP, " +
                    "SUM(MapQty * RATIO * WagePrice) MAN_DAY_INMAP,count(*) ITEM_COUNT, SUM(MapQty * tndPrice * ISNULL(BUDGET_RATIO, 100)/100) BUDGET, SUM(MapQty * RATIO * WagePrice * ISNULL(BUDGET_WAGE_RATIO, 100)/100) BUDGET_WAGE, " +
                    "SUM(MapQty * ITEM_UNIT_COST) + SUM(MapQty * MAN_PRICE) P_COST, SUM(tndQTY * ITEM_UNIT_PRICE) TND_COST FROM " +
                    "(SELECT pi.*, w.RATIO, w.PRICE, map.QTY MapQty, ISNULL(p.WAGE_MULTIPLIER, 0) AS WagePrice, it.ITEM_UNIT_PRICE AS tndPrice, it.ITEM_QUANTITY AS tndQTY FROM PLAN_ITEM pi LEFT OUTER JOIN " +
                    "TND_WAGE w ON pi.PLAN_ITEM_ID = w.PROJECT_ITEM_ID LEFT OUTER JOIN vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID LEFT JOIN TND_PROJECT p ON pi.PROJECT_ID = p.PROJECT_ID LEFT OUTER JOIN TND_PROJECT_ITEM it " +
                    "ON it.PROJECT_ITEM_ID = pi.PLAN_ITEM_ID WHERE it.project_id = @projectid) A  " +
                    "GROUP BY TYPE_CODE_1, TYPE_CODE_2)B ";
                logger.Info("sql = " + sql);
                lstTotalCost = context.Database.SqlQuery<DirectCost>(sql, new SqlParameter("projectid", projectid)).First();
            }
            return lstTotalCost;
        }

        #region 取得特定標單項目材料成本與預算
        //取得特定標單項目材料成本與預算
        public DirectCost getItemBudget(string projectid, string typeCode1, string typeCode2, string systemMain, string systemSub, string formName)
        {
            logger.Info("search plan item by 九宮格 =" + typeCode1 + "search plan item by 次九宮格 =" + typeCode2 + "search plan item by 主系統 =" + systemMain + "search plan item by 次系統 =" + systemSub + "search plan item by 採購名稱 =" + formName);
            DirectCost lstItemBudget = null;
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT SUM(map.QTY*pi.ITEM_UNIT_PRICE) AS ITEM_COST, SUM(map.QTY*tpi.ITEM_UNIT_PRICE*ISNULL(pi.BUDGET_RATIO, 100)/100) AS ITEM_BUDGET, SUM(map.QTY*w.RATIO*ISNULL(p.WAGE_MULTIPLIER, 0)*ISNULL(pi.BUDGET_WAGE_RATIO, 100)/100) AS ITEM_BUDGET_WAGE FROM PLAN_ITEM pi " +
                "LEFT JOIN TND_PROJECT_ITEM tpi ON pi.PLAN_ITEM_ID = tpi.PROJECT_ITEM_ID LEFT OUTER JOIN TND_WAGE w ON pi.PLAN_ITEM_ID = w.PROJECT_ITEM_ID LEFT JOIN TND_PROJECT p ON pi.PROJECT_ID = p.PROJECT_ID " +
                "LEFT JOIN vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID ";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectid));
            //採購項目
            if (null != formName && formName != "")
            {
                sql = sql + "right join (select distinct p.FORM_NAME + pii.PLAN_ITEM_ID as FORM_KEY, p.FORM_NAME, pii.PLAN_ITEM_ID FROM PLAN_SUP_INQUIRY p " +
                    "LEFT JOIN PLAN_SUP_INQUIRY_ITEM pii on p.INQUIRY_FORM_ID = pii.INQUIRY_FORM_ID WHERE p.FORM_NAME =@formName)A ON pi.PLAN_ITEM_ID = A.PLAN_ITEM_ID ";
                parameters.Add(new SqlParameter("formName", formName));
            }
            sql = sql + " WHERE pi.PROJECT_ID =@projectid ";
            //九宮格
            if (null != typeCode1 && typeCode1 != "")
            {
                sql = sql + "AND pi.TYPE_CODE_1 = @typeCode1 ";
                parameters.Add(new SqlParameter("typeCode1", typeCode1));
            }
            //次九宮格
            if (null != typeCode2 && typeCode2 != "")
            {
                sql = sql + "AND pi.TYPE_CODE_2 = @typeCode2 ";
                parameters.Add(new SqlParameter("typeCode2", typeCode2));
            }
            //主系統
            if (null != systemMain && systemMain != "")
            {
                sql = sql + "AND pi.SYSTEM_MAIN LIKE @systemMain ";
                parameters.Add(new SqlParameter("systemMain", "%" + systemMain + "%"));
            }
            //次系統
            if (null != systemSub && systemSub != "")
            {
                sql = sql + "AND pi.SYSTEM_SUB LIKE @systemSub ";
                parameters.Add(new SqlParameter("systemSub", "%" + systemSub + "%"));
            }

            using (var context = new topmepEntities())
            {
                logger.Debug("get plan item sql=" + sql);
                lstItemBudget = context.Database.SqlQuery<DirectCost>(sql, parameters.ToArray()).First();
            }
            return lstItemBudget;
        }
        #endregion

    }
    //採購詢價單資料提供作業
    public class PurchaseFormService : TnderProjectService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public PLAN_SUP_INQUIRY formInquiry = null;
        public List<PLAN_SUP_INQUIRY_ITEM> formInquiryItem = null;
        public Dictionary<string, COMPARASION_DATA_4PLAN> dirSupplierQuo = null;
        public PurchaseFormModel POFormData = null;
        public List<FIN_SUBJECT> ExpBudgetItem = null;
        public CashFlowModel cashFlowModel = new CashFlowModel();
        string sql4SupInqueryForm = @"
SELECT a.INQUIRY_FORM_ID, a.SUPPLIER_ID, a.FORM_NAME, SUM(b.ITEM_QTY* b.ITEM_UNIT_PRICE) AS TOTAL_PRICE, 
 ROW_NUMBER() OVER(ORDER BY a.INQUIRY_FORM_ID DESC) AS NO, ISNULL(a.STATUS, '有效') AS STATUS, ISNULL(a.ISWAGE, 'N') ISWAGE 
 FROM PLAN_SUP_INQUIRY a left JOIN PLAN_SUP_INQUIRY_ITEM b ON a.INQUIRY_FORM_ID = b.INQUIRY_FORM_ID 
 WHERE ISNULL(a.STATUS,'有效')=@status AND ISNULL(a.ISWAGE,'N')=@type 
 AND ISNULL(a.SUPPLIER_ID,'') !='' 
";

        #region 取得得標標單項目內容
        //取得標單品項資料
        public List<PlanItem4Map> getPlanItem(string checkEx, string projectid, string typeCode1, string typeCode2, string systemMain, string systemSub, string formName, string supplier, string delFlg, string inContractFlg)
        {

            logger.Info("search plan item by 九宮格 =" + typeCode1 + "search plan item by 次九宮格 =" + typeCode2 + "search plan item by 主系統 =" + systemMain + "search plan item by 次系統 =" + systemSub + "search plan item by 採購項目 =" + formName + "search plan item by 材料供應商 =" + supplier);
            List<topmeperp.Models.PlanItem4Map> lstItem = new List<PlanItem4Map>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT pi.*, map.QTY AS MAP_QTY FROM PLAN_ITEM pi LEFT JOIN vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID ";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectid));

            //採購項目
            if (null != formName && formName != "")
            {
                sql = sql + "right join (select distinct p.FORM_NAME + pii.PLAN_ITEM_ID as FORM_KEY, p.FORM_NAME, pii.PLAN_ITEM_ID FROM PLAN_SUP_INQUIRY p " +
                    "LEFT JOIN PLAN_SUP_INQUIRY_ITEM pii on p.INQUIRY_FORM_ID = pii.INQUIRY_FORM_ID WHERE p.FORM_NAME =@formName)A ON pi.PLAN_ITEM_ID = A.PLAN_ITEM_ID ";
                parameters.Add(new SqlParameter("formName", formName));
            }
            sql = sql + " WHERE pi.PROJECT_ID =@projectid ";
            //九宮格
            if (null != typeCode1 && typeCode1 != "")
            {
                sql = sql + "AND pi.TYPE_CODE_1 = @typeCode1 ";
                parameters.Add(new SqlParameter("typeCode1", typeCode1));
            }
            //次九宮格
            if (null != typeCode2 && typeCode2 != "")
            {
                sql = sql + "AND pi.TYPE_CODE_2 = @typeCode2 ";
                parameters.Add(new SqlParameter("typeCode2", typeCode2));
            }
            //主系統
            if (null != systemMain && systemMain != "")
            {
                sql = sql + "AND pi.SYSTEM_MAIN LIKE @systemMain ";
                parameters.Add(new SqlParameter("systemMain", "%" + systemMain + "%"));
            }
            //次系統
            if (null != systemSub && systemSub != "")
            {
                sql = sql + "AND pi.SYSTEM_SUB LIKE @systemSub ";
                parameters.Add(new SqlParameter("systemSub", "%" + systemSub + "%"));
            }
            //材料供應商
            if (null != supplier && supplier != "")
            {
                sql = sql + "AND pi.SUPPLIER_ID =@supplier ";
                parameters.Add(new SqlParameter("supplier", supplier));
            }
            //顯示未分類資料
            if (null != checkEx && checkEx != "")
            {
                sql = sql + "AND pi.TYPE_CODE_1 is null or pi.PROJECT_ID =@projectid AND pi.TYPE_CODE_1='' ";
            }
            //刪除註記
            if ("*" != delFlg)
            {
                sql = sql + "AND ISNULL(pi.DEL_FLAG,'N')=@delFlg ";
                parameters.Add(new SqlParameter("delFlg", delFlg));
            }

            //合約內品項標籤
            if ("*" != inContractFlg)
            {
                sql = sql + "AND ISNULL(IN_CONTRACT,'Y')=@inContract ";
                parameters.Add(new SqlParameter("inContract", inContractFlg));
            }

            sql = sql + "  ORDER BY EXCEL_ROW_ID;";
            using (var context = new topmepEntities())
            {
                logger.Debug("get plan item sql=" + sql);
                lstItem = context.Database.SqlQuery<PlanItem4Map>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get plan item count=" + lstItem.Count);
            return lstItem;
        }
        #endregion

        public PLAN_ITEM getPlanItem(string itemid)
        {
            logger.Debug("get plan item by id=" + itemid);
            PLAN_ITEM pitem = null;
            using (var context = new topmepEntities())
            {
                string sql = @"SELECT PLAN_ITEM_ID
                             ,PROJECT_ID
                             ,ITEM_ID
                             ,ITEM_DESC
                             ,ITEM_UNIT
                             ,ITEM_QUANTITY
                             ,ITEM_UNIT_PRICE
                             ,MAN_PRICE
                             ,ITEM_REMARK
                             ,TYPE_CODE_1
                             ,TYPE_CODE_2
                             ,SUB_TYPE_CODE
                             ,SYSTEM_MAIN
                             ,SYSTEM_SUB
                             ,MODIFY_USER_ID
                             ,MODIFY_DATE
                             ,CREATE_USER_ID
                             ,CREATE_DATE
                             ,EXCEL_ROW_ID
                             ,FORM_NAME
                             ,SUPPLIER_ID
                             ,BUDGET_RATIO
                             ,ITEM_FORM_QUANTITY
                             ,ITEM_UNIT_COST
                             ,TND_RATIO
                             ,MAN_FORM_NAME
                             ,MAN_SUPPLIER_ID
                             ,LEAD_TIME
                             ,DEL_FLAG
                             ,INQUIRY_FORM_ID
                             ,MAN_FORM_ID
                             ,BUDGET_WAGE_RATIO
                             ,ISNULL(IN_CONTRACT,'Y') IN_CONTRACT
                              FROM PLAN_ITEM WHERE PLAN_ITEM_ID=@itemid";
                //條件篩選
                pitem = context.PLAN_ITEM.SqlQuery(sql, new SqlParameter("itemid", itemid)).First();
            }
            return pitem;
        }

        public int updatePlanItem(PLAN_ITEM item)
        {
            int i = 0;
            if (null == item.PLAN_ITEM_ID || item.PLAN_ITEM_ID == "")
            {
                logger.Debug("add new plan item in porjectid=" + item.PROJECT_ID);
                item = getNewPlanItemID(item);
            }
            logger.Debug("plan item key=" + item.PLAN_ITEM_ID);
            using (var context = new topmepEntities())
            {
                try
                {
                    context.PLAN_ITEM.AddOrUpdate(item);
                    i = context.SaveChanges();
                }
                catch (Exception e)
                {
                    logger.Error("updatePlanItem  fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }
            }
            return i;
        }

        private PLAN_ITEM getNewPlanItemID(PLAN_ITEM item)
        {
            string sql = "SELECT MAX(CAST(SUBSTRING(PLAN_ITEM_ID,8,LEN(PLAN_ITEM_ID)) AS INT) +1) MaxSN, MAX(EXCEL_ROW_ID) + 1 as Row "
                + " FROM PLAN_ITEM WHERE PROJECT_ID = @projectid ; ";
            var parameters = new Dictionary<string, Object>();
            parameters.Add("projectid", item.PROJECT_ID);
            DataSet ds = ExecuteStoreQuery(sql, CommandType.Text, parameters);
            logger.Debug("sql=" + sql + "," + ds.Tables[0].Rows[0][0].ToString() + "," + ds.Tables[0].Rows[0][1].ToString());
            int longMaxExcel = 1;
            int longMaxItem = 1;
            if (DBNull.Value != ds.Tables[0].Rows[0][0])
            {
                longMaxItem = int.Parse(ds.Tables[0].Rows[0][0].ToString());
                longMaxExcel = int.Parse(ds.Tables[0].Rows[0][1].ToString());
            }
            logger.Debug("new plan item id=" + longMaxItem + ",ExcelRowID=" + longMaxExcel);
            item.PLAN_ITEM_ID = item.PROJECT_ID + "-" + longMaxItem;
            //新品項不會有Excel Row_id
            if (null == item.EXCEL_ROW_ID || item.EXCEL_ROW_ID == 0)
            {
                item.EXCEL_ROW_ID = longMaxExcel;
            }
            return item;
        }

        //於現有品項下方新增一筆資料
        public int addPlanItemAfter(PLAN_ITEM item)
        {
            string sql = "UPDATE PLAN_ITEM SET EXCEL_ROW_ID=EXCEL_ROW_ID+1 WHERE PROJECT_ID = @projectid AND EXCEL_ROW_ID> @ExcelRowId ";

            using (var db = new topmepEntities())
            {
                logger.Debug("add exce rowid sql=" + sql + ",projectid=" + item.PROJECT_ID + ",ExcelRowI=" + item.EXCEL_ROW_ID);
                db.Database.ExecuteSqlCommand(sql, new SqlParameter("projectid", item.PROJECT_ID), new SqlParameter("ExcelRowId", item.EXCEL_ROW_ID));
            }
            item.PLAN_ITEM_ID = "";
            item.ITEM_UNIT_COST = null;
            item.EXCEL_ROW_ID = item.EXCEL_ROW_ID + 1;
            return updatePlanItem(item);
        }
        //將Plan Item 註記刪除
        public int changePlanItem(string itemid, string delFlag)
        {
            string sql = "UPDATE PLAN_ITEM SET DEL_FLAG=@delFlag WHERE PLAN_ITEM_ID = @itemid";
            int i = 0;
            using (var db = new topmepEntities())
            {
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("itemid", itemid));
                parameters.Add(new SqlParameter("delFlag", delFlag));
                logger.Info("Update PLAN_ITEM FLAG=" + sql + ",itemid=" + itemid + ",delFlag=" + delFlag);
                i = db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            }
            return i;
        }

        //批次產生空白表單
        public int createPlanEmptyForm(string projectid, SYS_USER loginUser)
        {
            int i = 0;
            int i2 = 0;
            using (var context = new topmepEntities())
            {
                //0.清除所有空白詢價單樣板//僅刪除材料之空白詢價單
                string sql = "DELETE FROM PLAN_SUP_INQUIRY_ITEM WHERE INQUIRY_FORM_ID IN (SELECT INQUIRY_FORM_ID FROM PLAN_SUP_INQUIRY WHERE SUPPLIER_ID IS NULL AND PROJECT_ID=@projectid AND ISNULL(ISWAGE,'N')='N');";
                i2 = context.Database.ExecuteSqlCommand(sql, new SqlParameter("projectid", projectid));
                logger.Info("delete template inquiry form item  by porjectid=" + projectid + ",result=" + i2);
                sql = "DELETE FROM PLAN_SUP_INQUIRY WHERE SUPPLIER_ID IS NULL AND PROJECT_ID=@projectid AND ISNULL(ISWAGE,'N')='N'; ";
                i2 = context.Database.ExecuteSqlCommand(sql, new SqlParameter("projectid", projectid));
                logger.Info("delete template inquiry form  by porjectid=" + projectid + ",result=" + i2);

                //1.依據專案取得九宮格次九宮格分類.
                sql = "SELECT DISTINCT isnull(TYPE_CODE_1,'未分類') TYPE_CODE_1," +
                   "(SELECT TYPE_DESC FROM REF_TYPE_MAIN m WHERE m.TYPE_CODE_1 + m.TYPE_CODE_2 = p.TYPE_CODE_1) as TYPE_CODE_1_NAME, " +
                   "isnull(TYPE_CODE_2,'未分類') TYPE_CODE_2," +
                   "(SELECT TYPE_DESC FROM REF_TYPE_SUB sub WHERE sub.TYPE_CODE_ID = p.TYPE_CODE_1 AND sub.SUB_TYPE_CODE = p.TYPE_CODE_2) as TYPE_CODE_2_NAME " +
                   "FROM PLAN_ITEM p WHERE PROJECT_ID = @projectid ORDER BY TYPE_CODE_1 ,Type_CODE_2; ";

                List<TYPE_CODE_INDEX> lstType = context.Database.SqlQuery<TYPE_CODE_INDEX>(sql, new SqlParameter("projectid", projectid)).ToList();
                logger.Debug("get type index count=" + lstType.Count);
                foreach (TYPE_CODE_INDEX idx in lstType)
                {
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("projectid", projectid));
                    sql = "SELECT * FROM PLAN_ITEM WHERE PROJECT_ID = @projectid ";
                    if (idx.TYPE_CODE_1 == "未分類")
                    {
                        sql = sql + "AND TYPE_CODE_1 is null ";
                    }
                    else
                    {
                        sql = sql + "AND TYPE_CODE_1=@typecode1 ";
                        parameters.Add(new SqlParameter("typecode1", idx.TYPE_CODE_1));
                    }

                    if (idx.TYPE_CODE_2 == "未分類")
                    {
                        sql = sql + "AND TYPE_CODE_2 is null ";
                    }
                    else
                    {
                        sql = sql + "AND TYPE_CODE_2=@typecode2 ";
                        parameters.Add(new SqlParameter("typecode2", idx.TYPE_CODE_2));
                    }
                    //2.依據分類取得詢價單項次
                    List<PLAN_ITEM> lstPlanItem = context.PLAN_ITEM.SqlQuery(sql, parameters.ToArray()).ToList();
                    logger.Debug("get plan item count=" + lstPlanItem.Count + ", by typecode1=" + idx.TYPE_CODE_1 + ",typeCode2=" + idx.TYPE_CODE_2);
                    string[] itemId = new string[lstPlanItem.Count];
                    int j = 0;
                    foreach (PLAN_ITEM item in lstPlanItem)
                    {
                        itemId[j] = item.PLAN_ITEM_ID;
                        j++;
                    }
                    //3.建立詢價單基本資料
                    PLAN_SUP_INQUIRY f = new PLAN_SUP_INQUIRY();
                    if (idx.TYPE_CODE_1 == "未分類")
                    {
                        f.FORM_NAME = "未分類";
                    }
                    else
                    {
                        f.FORM_NAME = idx.TYPE_CODE_1_NAME;
                    }

                    if (idx.TYPE_CODE_2 != "未分類")
                    {
                        f.FORM_NAME = f.FORM_NAME + "-" + idx.TYPE_CODE_2_NAME;
                    }
                    f.FORM_NAME = "(" + idx.TYPE_CODE_1 + "," + idx.TYPE_CODE_2 + ")" + f.FORM_NAME;
                    f.PROJECT_ID = projectid;
                    f.CREATE_ID = loginUser.USER_ID;
                    f.CREATE_DATE = DateTime.Now;
                    f.OWNER_NAME = loginUser.USER_NAME;
                    f.OWNER_EMAIL = loginUser.EMAIL;
                    f.OWNER_TEL = loginUser.TEL;
                    f.OWNER_FAX = loginUser.FAX;
                    //4.建立表單
                    string fid = newPlanForm(f, itemId);
                    logger.Info("create template form:" + fid);
                    i++;
                }
            }
            logger.Info("create form count" + i);
            return i;
        }

        public string newPlanForm(PLAN_SUP_INQUIRY form, string[] lstItemId)
        {
            //1.建立詢價單價單樣本
            logger.Info("create new plan form ");
            string sno_key = "PP";
            SerialKeyService snoservice = new SerialKeyService();
            form.INQUIRY_FORM_ID = snoservice.getSerialKey(sno_key);
            logger.Info("new plan form =" + form.ToString());
            using (var context = new topmepEntities())
            {
                context.PLAN_SUP_INQUIRY.Add(form);
                int i = context.SaveChanges();
                logger.Debug("Add form=" + i);
                logger.Info("plan form id = " + form.INQUIRY_FORM_ID);
                //if (i > 0) { status = true; };
                List<topmeperp.Models.PLAN_SUP_INQUIRY_ITEM> lstItem = new List<PLAN_SUP_INQUIRY_ITEM>();
                string ItemId = "";
                for (i = 0; i < lstItemId.Count(); i++)
                {
                    if (i < lstItemId.Count() - 1)
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'" + ",";
                    }
                    else
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'";
                    }
                }

                string sql = "INSERT INTO PLAN_SUP_INQUIRY_ITEM (INQUIRY_FORM_ID, PLAN_ITEM_ID, TYPE_CODE, "
                    + "SUB_TYPE_CODE, ITEM_DESC, ITEM_UNIT, ITEM_QTY, ITEM_UNIT_PRICE, ITEM_REMARK,ITEM_ID) "
                    + "SELECT '" + form.INQUIRY_FORM_ID + "' as INQUIRY_FORM_ID, PLAN_ITEM_ID, TYPE_CODE_1 AS TYPE_CODE, "
                    + "TYPE_CODE_2 AS SUB_TYPE_CODE, ITEM_DESC, ITEM_UNIT, map.QTY, ITEM_UNIT_COST, ITEM_REMARK,pi.ITEM_ID ITEM_ID "
                    + "FROM PLAN_ITEM pi LEFT OUTER JOIN vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID where PLAN_ITEM_ID IN (" + ItemId + ")";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return form.INQUIRY_FORM_ID;
            }
        }
        //取得採購詢價單
        public void getInqueryForm(string formid)
        {
            logger.Info("get form : formid=" + formid);
            using (var context = new topmepEntities())
            {
                //取得詢價單檔頭資訊
                string sql = "SELECT INQUIRY_FORM_ID,PROJECT_ID,FORM_NAME,OWNER_NAME,OWNER_TEL "
                    + ",OWNER_EMAIL, OWNER_FAX, SUPPLIER_ID, CONTACT_NAME, CONTACT_EMAIL "
                    + ",DUEDATE, REF_ID, CREATE_ID, CREATE_DATE, MODIFY_ID"
                    + ",MODIFY_DATE,ISNULL(STATUS,'有效') as STATUS,ISNULL(ISWAGE,'N') as ISWAGE "
                    + "FROM PLAN_SUP_INQUIRY WHERE INQUIRY_FORM_ID = @formid";
                formInquiry = context.PLAN_SUP_INQUIRY.SqlQuery(sql, new SqlParameter("formid", formid)).First();
                //取得詢價單明細
                formInquiryItem = context.PLAN_SUP_INQUIRY_ITEM.SqlQuery("SELECT i.[INQUIRY_ITEM_ID],i.[INQUIRY_FORM_ID]" +
                    ", i.[PLAN_ITEM_ID], i.[TYPE_CODE], i.[SUB_TYPE_CODE], pi.[ITEM_ID], i.[ITEM_DESC], i.[ITEM_UNIT] "
                    + " , i.[ITEM_QTY],i.[ITEM_UNIT_PRICE], i.[ITEM_QTY_ORG] , i.[ITEM_UNITPRICE_ORG], i.ITEM_REMARK "
                    + " , i.[MODIFY_ID], i.[MODIFY_DATE], i.[WAGE_PRICE]  "
                    + "FROM PLAN_SUP_INQUIRY_ITEM i LEFT OUTER JOIN  PLAN_ITEM pi on i.PLAN_ITEM_ID = pi.PLAN_ITEM_ID "
                    + "WHERE i.INQUIRY_FORM_ID=@formid ORDER BY pi.EXCEL_ROW_ID", new SqlParameter("formid", formid)).ToList();
                logger.Debug("get form item count:" + formInquiryItem.Count);
            }
        }
        int i = 0;
        // 取得採購詢價單預算金額
        public List<COMPARASION_DATA_4PLAN> getBudgetForComparison(string projectid, string formname)
        {
            List<COMPARASION_DATA_4PLAN> budget = new List<COMPARASION_DATA_4PLAN>();
            string[] eachname = formname.Split(',');
            string ItemId = "";
            for (i = 0; i < eachname.Count(); i++)
            {
                if (i < eachname.Count() - 1)
                {
                    ItemId = ItemId + "'" + eachname[i] + "'" + ",";
                }
                else
                {
                    ItemId = ItemId + "'" + eachname[i] + "'";
                }
            }
            string sql = "SELECT FORM_NAME, BUDGET_AMOUNT AS BAmount " +
                "FROM PLAN_BUDGET WHERE PROJECT_ID = @pid AND FORM_NAME IN (" + ItemId + ") ";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("pid", projectid));
            using (var context = new topmepEntities())
            {
                logger.Debug("get sql=" + sql);
                budget = context.Database.SqlQuery<COMPARASION_DATA_4PLAN>(sql, parameters.ToArray()).ToList();
            }
            return budget;
        }
        public string zipAllTemplate4Download(string projectid)
        {
            //1.取得專案所有空白詢價單
            List<PLAN_SUP_INQUIRY> lstTemplate = getFormTemplateByProject(projectid);
            ZipFileCreator zipTool = new ZipFileCreator();
            //2.設定暫存目錄
            string tempFolder = ContextService.strUploadPath + "\\" + projectid + "\\" + ContextService.quotesFolder + "\\Temp\\";
            ZipFileCreator.DelDirectory(tempFolder);
            ZipFileCreator.CreateDirectory(tempFolder);
            //3.批次產生空白詢價單
            PurchaseFormtoExcel poi = new PurchaseFormtoExcel();
            TND_PROJECT p = getProjectById(projectid);
            foreach (PLAN_SUP_INQUIRY f in lstTemplate)
            {
                getInqueryForm(f.INQUIRY_FORM_ID);
                ZipFileCreator.CreateDirectory(tempFolder + formInquiry.FORM_NAME);
                string fileLocation = poi.exportExcel4po(formInquiry, formInquiryItem, true, false);
                logger.Debug("temp file=" + fileLocation);
            }
            //4.Zip all file
            return zipTool.ZipDirectory(tempFolder);
            //return zipTool.ZipFiles(tempFolder, null, p.PROJECT_NAME);
        }
        //取得採購詢價單樣板(供應商欄位為0)
        public List<PLAN_SUP_INQUIRY> getFormTemplateByProject(string projectid)
        {
            logger.Info("get purchase template by projectid=" + projectid);
            List<PLAN_SUP_INQUIRY> lst = new List<PLAN_SUP_INQUIRY>();
            using (var context = new topmepEntities())
            {
                //取得詢價單樣本資訊
                string sql = "SELECT INQUIRY_FORM_ID, PROJECT_ID,FORM_NAME,OWNER_NAME,OWNER_TEL,OWNER_EMAIL "
                    + ",OWNER_FAX,SUPPLIER_ID,CONTACT_NAME,CONTACT_EMAIL,DUEDATE,REF_ID,CREATE_ID,CREATE_DATE "
                    + ",MODIFY_ID,MODIFY_DATE,ISNULL(STATUS,'有效') STATUS, ISNULL(ISWAGE,'N') ISWAGE "
                    + "FROM PLAN_SUP_INQUIRY WHERE SUPPLIER_ID IS NULL AND PROJECT_ID =@projectid ORDER BY INQUIRY_FORM_ID DESC";
                lst = context.PLAN_SUP_INQUIRY.SqlQuery(sql, new SqlParameter("projectid", projectid)).ToList();
            }
            return lst;
        }
        //採發階段發包分項與預算 (材料預算)
        public PurchaseFormModel getInquiryWithBudget(TND_PROJECT project, string status)
        {
            POFormData = new PurchaseFormModel();
            getBudgetSummary(project);
            POFormData.materialTemplateWithBudget = getTemplateRefBudget(project, "N", status);
            POFormData.wageTemplateWithBudget = getTemplateRefBudget(project, "Y", status);
            return POFormData;
        }
        //取得詢價單樣本與分項預算
        public IEnumerable<PURCHASE_ORDER> getTemplateRefBudget(TND_PROJECT project, string iswage, string status)
        {
            logger.Info("get purchase template by projectid=" + project.PROJECT_ID);
            List<PURCHASE_ORDER> lst = new List<PURCHASE_ORDER>();
            string sql = "";
            decimal wageunitprice = 2500;
            using (var context = new topmepEntities())
            {
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", project.PROJECT_ID));
                parameters.Add(new SqlParameter("status", status));
                if (iswage == "N")
                {
                    //取得詢價單樣本資訊 - 材料預算-圖算數量*報價標單(Project_item)單價 * 預算折扣比率
                    sql = @"
SELECT tmp.*,CountPO, 
(SELECT SUM(v.QTY * tpi.ITEM_UNIT_PRICE * it.BUDGET_RATIO/100 ) as BudgetAmount 
FROM PLAN_ITEM it LEFT JOIN vw_MAP_MATERLIALIST v ON it.PLAN_ITEM_ID = v.PROJECT_ITEM_ID  
LEFT JOIN TND_PROJECT_ITEM tpi ON it.PLAN_ITEM_ID = tpi.PROJECT_ITEM_ID
 WHERE it.PLAN_ITEM_ID in (SELECT  iit.PLAN_ITEM_ID FROM PLAN_SUP_INQUIRY_ITEM iit 
 WHERE iit.INQUIRY_FORM_ID = tmp.INQUIRY_FORM_ID)) AS BudgetAmount 
 FROM(
 SELECT * FROM PLAN_SUP_INQUIRY 
 WHERE ISNULL(SUPPLIER_ID,'') ='' AND PROJECT_ID = @projectid AND ISNULL(STATUS,'有效')=@status 
 AND ISNULL(ISWAGE,'N')='N') tmp 
 LEFT OUTER JOIN (
 SELECT COUNT(*) CountPO, FORM_NAME, PROJECT_ID 
 FROM  PLAN_SUP_INQUIRY 
 WHERE SUPPLIER_ID IS NOT Null 
 GROUP BY FORM_NAME, PROJECT_ID
 ) Quo ON Quo.PROJECT_ID = tmp.PROJECT_ID AND Quo.FORM_NAME = tmp.FORM_NAME ";
                }
                else
                {
                    // 取得詢價單樣本資訊 - 工資預算 - 圖算數量 * 工資單(預設2500) * 工率*  預算折扣比率
                    sql = @"
SELECT tmp.*,CountPO,(
SELECT SUM(v.QTY * w.RATIO * it.BUDGET_WAGE_RATIO / 100 * @wageunitprice) as BudgetAmount 
FROM PLAN_ITEM it 
LEFT JOIN vw_MAP_MATERLIALIST v 
ON it.PLAN_ITEM_ID = v.PROJECT_ITEM_ID LEFT JOIN TND_WAGE w 
ON it.PLAN_ITEM_ID = w.PROJECT_ITEM_ID WHERE it.PLAN_ITEM_ID in (
SELECT  iit.PLAN_ITEM_ID 
FROM PLAN_SUP_INQUIRY_ITEM iit 
WHERE iit.INQUIRY_FORM_ID = tmp.INQUIRY_FORM_ID)) AS BudgetAmount 
FROM(
SELECT * FROM PLAN_SUP_INQUIRY 
WHERE ISNULL(SUPPLIER_ID,'') = '' 
AND PROJECT_ID = @projectid 
AND ISNULL(STATUS, '有效')=@status AND ISNULL(ISWAGE, 'N') = 'Y'
) tmp LEFT OUTER JOIN (
SELECT COUNT(*) CountPO, FORM_NAME, PROJECT_ID FROM  PLAN_SUP_INQUIRY 
WHERE SUPPLIER_ID IS NOT Null GROUP BY FORM_NAME, PROJECT_ID
) Quo ON Quo.PROJECT_ID = tmp.PROJECT_ID AND Quo.FORM_NAME = tmp.FORM_NAME ;
";
                    if (null != project.WAGE_MULTIPLIER)
                    {
                        wageunitprice = (decimal)project.WAGE_MULTIPLIER;
                    }
                    parameters.Add(new SqlParameter("wageunitprice", wageunitprice));
                }

                logger.Debug("sql=" + sql + ",projectId=" + project.PROJECT_ID);
                lst = context.Database.SqlQuery<PURCHASE_ORDER>(sql, parameters.ToArray()).ToList();
            }
            return lst;
        }
        //取得預算總價
        public void getBudgetSummary(TND_PROJECT project)
        {
            string sql = "SELECT SUM(mBudget) Material_Budget,SUM(wBudget) Wage_Budget FROM ("
                + "SELECT pi.PLAN_ITEM_ID,pi.PROJECT_ID,pi.ITEM_ID,pi.ITEM_DESC,pi.ITEM_QUANTITY,map.QTY mapQty, pi.ITEM_UNIT,"
                + "pi.ITEM_UNIT_PRICE SellProice, pji.ITEM_UNIT_PRICE Cost, isNull(pi.BUDGET_RATIO, 100) BUDGET_RATIO,"
                + "(map.QTY * pji.ITEM_UNIT_PRICE * isNull(pi.BUDGET_RATIO, 100) / 100) mBudget,"
                + "isNull(pi.BUDGET_WAGE_RATIO, 100) BUDGET_WAGE_RATIO,isnull(w.RATIO, 0) wRatio,"
                + "(@wageunitprice * ISNULL(map.QTY, 0) * isNull(pi.BUDGET_WAGE_RATIO, 100) * isnull(w.RATIO, 0) / 100) wBudget "
                + "FROM PLAN_ITEM pi LEFT OUTER JOIN TND_PROJECT_ITEM pji on pi.PLAN_ITEM_ID = pji.PROJECT_ITEM_ID "
                + "LEFT OUTER JOIN  vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID "
                + "LEFT OUTER JOIN TND_WAGE w ON pi.PLAN_ITEM_ID = w.PROJECT_ITEM_ID "
                + "WHERE pi.PROJECT_ID = @projectid) A; ";
            using (var context = new topmepEntities())
            {
                var parameters = new List<SqlParameter>();
                //設定專案預算資料
                parameters.Add(new SqlParameter("projectid", project.PROJECT_ID));
                //專案工資若未設定則以2500 計算
                if (null == project.WAGE_MULTIPLIER)
                {
                    project.WAGE_MULTIPLIER = 2500;
                }
                parameters.Add(new SqlParameter("wageunitprice", project.WAGE_MULTIPLIER));
                logger.Debug("sql=" + sql + ",projectId=" + project.PROJECT_ID);
                POFormData.BudgetSummary = context.Database.SqlQuery<BUDGET_SUMMANY>(sql, parameters.ToArray()).First();
            }
        }
        //取得廠商合約報價單
        public List<PlanSupplierFormFunction> getContractForm(string projectid, string status, string type, string formname)
        {
            List<PlanSupplierFormFunction> lst = new List<PlanSupplierFormFunction>();
            string sql = sql4SupInqueryForm + " AND a.INQUIRY_FORM_ID IN (SELECT C.INQUIRY_FORM_ID FROM PLAN_ITEM2_SUP_INQUIRY C WHERE PROJECT_ID=@projectId) ";
            sql = sql + "AND a.FORM_NAME LIKE @formname ";
            sql = sql + " GROUP BY a.INQUIRY_FORM_ID, a.SUPPLIER_ID, a.FORM_NAME, a.PROJECT_ID, a.STATUS, a.ISWAGE HAVING  a.SUPPLIER_ID IS NOT NULL " +
        "AND a.PROJECT_ID =@projectid ORDER BY a.INQUIRY_FORM_ID DESC, a.FORM_NAME ;";
            logger.Info("sql4SupInqueryForm=" + sql);
            var parameters = new List<SqlParameter>();
            //設定專案編號資料
            parameters.Add(new SqlParameter("projectid", projectid));
            //設定詢價單是否有效
            parameters.Add(new SqlParameter("status", status));
            //設定詢價單為工資或材料
            parameters.Add(new SqlParameter("type", type));
            //詢價單名稱條件
            parameters.Add(new SqlParameter("formname", "%" + formname + "%"));
            using (var context = new topmepEntities())
            {
                lst = context.Database.SqlQuery<PlanSupplierFormFunction>(sql, parameters.ToArray()).ToList();
                logger.Info("get plan supplier form function count:" + lst.Count);
            }
            return lst;
        }
        /// <summary>
        /// 取消廠商合約報價單
        /// </summary>
        /// <returns></returns>
        public int cancelContractForm(string projectId, string formId)
        {
            logger.Info("delete contract form by proejct id=" + projectId + ",formId=" + formId);
            //1.刪除PLAN_ITEM2_SUP_INQUIRY
            string sql = "DELETE PLAN_ITEM2_SUP_INQUIRY WHERE PROJECT_ID = @projectId AND INQUIRY_FORM_ID = @formId; ";
            //2.更新PLAN_ITEM 資料
            string sqlUpdatePlanItem = @"UPDATE PLAN_ITEM SET
                FORM_NAME=NULL,SUPPLIER_ID=NULL,ITEM_FORM_QUANTITY=NULL,ITEM_UNIT_COST=NULL,INQUIRY_FORM_ID=NULL
                WHERE INQUIRY_FORM_ID=@formId AND PROJECT_ID=@projectId;
                UPDATE PLAN_ITEM SET
                MAN_FORM_ID=NULL,MAN_FORM_NAME=NULL,MAN_SUPPLIER_ID=NULL
                WHERE MAN_FORM_ID=@formId AND PROJECT_ID=@projectId;";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectId", projectId));
            parameters.Add(new SqlParameter("formId", formId));
            logger.Debug("Cance Form:" + sql + sqlUpdatePlanItem);
            using (var context = new topmepEntities())
            {
                i = context.Database.ExecuteSqlCommand(sql + sqlUpdatePlanItem, parameters.ToArray());
            }
            return i;
        }
        public List<PlanSupplierFormFunction> getFormByProject(string projectid, string _status, string _type, string formname)
        {
            string status = "有效";
            if (null != _status && _status != "*")
            {
                status = _status;
            }
            string type = "N";
            if (null != _type && _type != "*")
            {
                type = _type;
            }
            List<PlanSupplierFormFunction> lst = new List<PlanSupplierFormFunction>();
            string sql = sql4SupInqueryForm;
            var parameters = new List<SqlParameter>();
            //設定專案編號資料
            parameters.Add(new SqlParameter("projectid", projectid));
            //設定詢價單是否有效
            parameters.Add(new SqlParameter("status", status));
            //設定詢價單為工資或材料
            parameters.Add(new SqlParameter("type", type));
            //詢價單名稱條件
            if (null != formname && formname != "")
            {
                sql = sql + "AND a.FORM_NAME LIKE @formname ";
                parameters.Add(new SqlParameter("formname", "%" + formname + "%"));
            }
            sql = sql + " GROUP BY a.INQUIRY_FORM_ID, a.SUPPLIER_ID, a.FORM_NAME, a.PROJECT_ID, a.STATUS, a.ISWAGE HAVING  a.SUPPLIER_ID IS NOT NULL " +
                    "AND a.PROJECT_ID =@projectid ORDER BY a.INQUIRY_FORM_ID DESC, a.FORM_NAME ;";

            logger.Info("sql=" + sql);
            using (var context = new topmepEntities())
            {
                lst = context.Database.SqlQuery<PlanSupplierFormFunction>(sql, parameters.ToArray()).ToList();
                logger.Info("get plan supplier form function count:" + lst.Count);
            }
            return lst;
        }

        //取得尚未發包之分項詢價單資料(供應商欄位為0)
        public List<PURCHASE_ORDER> getFormTempOutOfContractByProject(string projectid)
        {
            logger.Info("get purchase template out of contract by projectid=" + projectid);
            List<PURCHASE_ORDER> lst = new List<PURCHASE_ORDER>();
            using (var context = new topmepEntities())
            {
                //取得詢價單樣本資訊
                string sql = "SELECT tmp.*,CountPO, Bargain, paymentFrequency, PAYMENT_TERMS, ContractId, Supplier, MATERIAL_BRAND, CONTRACT_PRODUCTION, " +
                       "CONVERT(char(10),DELIVERY_DATE, 111) AS DELIVERY_DATE, ConRemark, ROW_NUMBER() OVER(ORDER BY tmp.INQUIRY_FORM_ID) AS NO " +
                       "FROM(SELECT * FROM PLAN_SUP_INQUIRY WHERE SUPPLIER_ID is Null AND PROJECT_ID = @projectid AND ISNULL(STATUS, '有效') = '有效' AND ISNULL(ISWAGE, 'N') = 'N') tmp " +
                       "LEFT OUTER JOIN (SELECT COUNT(*) CountPO, FORM_NAME, PROJECT_ID FROM  PLAN_SUP_INQUIRY WHERE SUPPLIER_ID IS NOT Null GROUP BY FORM_NAME, PROJECT_ID) Quo " +
                       "ON Quo.PROJECT_ID = tmp.PROJECT_ID AND Quo.FORM_NAME = tmp.FORM_NAME LEFT OUTER JOIN (SELECT p.FORM_NAME AS Bargain, p.SUPPLIER_ID AS Supplier, IIF(ppt.PAYMENT_FREQUENCY = 'O', ppt.DATE_1, ppt.DATE_3) AS paymentFrequency, " +
                       "ppt.PAYMENT_TERMS, p.INQUIRY_FORM_ID AS ContractId, pcp.MATERIAL_BRAND, pcp.CONTRACT_PRODUCTION, pcp.DELIVERY_DATE, pcp.REMARK AS ConRemark FROM PLAN_ITEM p " +
                       "LEFT JOIN PLAN_PAYMENT_TERMS ppt ON p.INQUIRY_FORM_ID = ppt.CONTRACT_ID LEFT JOIN PLAN_CONTRACT_PROCESS pcp ON p.INQUIRY_FORM_ID = pcp.CONTRACT_ID WHERE p.PROJECT_ID = @projectid " +
                       "AND p.FORM_NAME IS NOT NULL GROUP BY p.FORM_NAME, p.SUPPLIER_ID, IIF(ppt.PAYMENT_FREQUENCY = 'O', ppt.DATE_1, ppt.DATE_3), ppt.PAYMENT_TERMS, p.INQUIRY_FORM_ID, " +
                       "pcp.MATERIAL_BRAND, pcp.CONTRACT_PRODUCTION, pcp.DELIVERY_DATE, pcp.REMARK)Con ON tmp.FORM_NAME = Con.Bargain UNION " +
                       "SELECT tmp.*,CountPO, Bargain, paymentFrequency, PAYMENT_TERMS, ContractId, Supplier, MATERIAL_BRAND, CONTRACT_PRODUCTION, " +
                       "CONVERT(char(10),DELIVERY_DATE, 111) AS DELIVERY_DATE, ConRemark, ROW_NUMBER() OVER(ORDER BY tmp.INQUIRY_FORM_ID) AS NO " +
                       "FROM(SELECT * FROM PLAN_SUP_INQUIRY WHERE SUPPLIER_ID is Null AND PROJECT_ID = @projectid AND ISNULL(STATUS, '有效') = '有效' AND ISNULL(ISWAGE, 'N') = 'Y') tmp " +
                       "LEFT OUTER JOIN (SELECT COUNT(*) CountPO, FORM_NAME, PROJECT_ID FROM  PLAN_SUP_INQUIRY WHERE SUPPLIER_ID IS NOT NULL GROUP BY FORM_NAME, PROJECT_ID) Quo " +
                       "ON Quo.PROJECT_ID = tmp.PROJECT_ID AND Quo.FORM_NAME = tmp.FORM_NAME  LEFT OUTER JOIN (SELECT p.MAN_FORM_NAME AS Bargain, p.MAN_SUPPLIER_ID AS Supplier, " +
                       "IIF(ppt.PAYMENT_FREQUENCY = 'O', ppt.DATE_1, ppt.DATE_3) AS paymentFrequency, ppt.PAYMENT_TERMS, p.MAN_FORM_ID AS ContractId, pcp.MATERIAL_BRAND, pcp.CONTRACT_PRODUCTION, " +
                       "pcp.DELIVERY_DATE, pcp.REMARK AS ConRemark FROM PLAN_ITEM p LEFT JOIN PLAN_PAYMENT_TERMS ppt ON p.MAN_FORM_ID = ppt.CONTRACT_ID LEFT JOIN " +
                       "PLAN_CONTRACT_PROCESS pcp ON p.MAN_FORM_ID = pcp.CONTRACT_ID WHERE p.PROJECT_ID = @projectid AND p.MAN_FORM_NAME IS NOT NULL " +
                       "GROUP BY p.MAN_FORM_NAME, p.MAN_SUPPLIER_ID, IIF(ppt.PAYMENT_FREQUENCY = 'O', ppt.DATE_1, ppt.DATE_3), ppt.PAYMENT_TERMS, p.MAN_FORM_ID, " +
                       "pcp.MATERIAL_BRAND, pcp.CONTRACT_PRODUCTION, pcp.DELIVERY_DATE, pcp.REMARK)Con ON tmp.FORM_NAME = Con.Bargain";

                logger.Info("sql =" + sql);
                lst = context.Database.SqlQuery<PURCHASE_ORDER>(sql, new SqlParameter("projectid", projectid)).ToList();
            }
            return lst;
        }

        public int delAllContractByProject(string projectId)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete all contract by proejct id=" + projectId);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_CONTRACT_PROCESS WHERE PROJECT_ID=@projectid AND CONTRACT_ID <> @projectid ", new SqlParameter("@projectid", projectId));
            }
            logger.Debug("delete contract count=" + i);
            return i;
        }
        //新增合約製作狀態
        public int AddContractProcess(string projectId, List<PLAN_CONTRACT_PROCESS> lstItem)
        {
            logger.Info("Add plan contract process by project id =" + projectId);
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    //將contract process寫入 
                    foreach (PLAN_CONTRACT_PROCESS item in lstItem)
                    {
                        PLAN_CONTRACT_PROCESS existItem = context.PLAN_CONTRACT_PROCESS.Find(item.CONTRACT_ID);
                        if (null == existItem)
                        {
                            existItem = new PLAN_CONTRACT_PROCESS();
                        }
                        logger.Debug("item contract id=" + item.CONTRACT_ID);
                        if (item.CONTRACT_ID != null)
                        {
                            existItem.CONTRACT_ID = item.CONTRACT_ID;
                            existItem.PROJECT_ID = projectId;
                            existItem.CONTRACT_PRODUCTION = item.CONTRACT_PRODUCTION;
                            existItem.DELIVERY_DATE = item.DELIVERY_DATE;
                            existItem.MATERIAL_BRAND = item.MATERIAL_BRAND;
                            existItem.REMARK = item.REMARK;
                            existItem.CREATE_ID = item.CREATE_ID;
                            existItem.CREATE_DATE = DateTime.Now;
                            context.PLAN_CONTRACT_PROCESS.AddOrUpdate(existItem);
                        }
                    }
                    j = context.SaveChanges();

                    logger.Debug("Add plan contract process item =" + j);
                }
                catch (Exception e)
                {
                    logger.Error("update new plan supplier form id fail:" + e.Message);
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }
            }
            return j;
        }
        public int addFormName(List<PLAN_SUP_INQUIRY> lstItem)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    logger.Info(" No. of plan form to refresh  = " + lstItem.Count);
                    //2.將plan form資料寫入 
                    foreach (PLAN_SUP_INQUIRY item in lstItem)
                    {
                        PLAN_SUP_INQUIRY existItem = null;
                        logger.Debug("plan form id=" + item.INQUIRY_FORM_ID);
                        if (item.INQUIRY_FORM_ID != null)
                        {
                            existItem = context.PLAN_SUP_INQUIRY.Find(item.INQUIRY_FORM_ID);
                        }
                        else
                        {
                            var parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("formid", item.INQUIRY_FORM_ID));
                            string sql = "SELECT * FROM PLAN_SUP_INQUIRY WHERE INQUIRY_FORM_ID=@formid";
                            logger.Info(sql + " ;" + item.INQUIRY_FORM_ID);
                            PLAN_SUP_INQUIRY excelItem = context.PLAN_SUP_INQUIRY.SqlQuery(sql, parameters.ToArray()).First();
                            existItem = context.PLAN_SUP_INQUIRY.Find(excelItem.INQUIRY_FORM_ID);

                        }
                        logger.Debug("find exist item=" + existItem.PROJECT_ID + " ;" + existItem.INQUIRY_FORM_ID);
                        existItem.FORM_NAME = item.FORM_NAME;
                        context.PLAN_SUP_INQUIRY.AddOrUpdate(existItem);
                    }
                    i = context.SaveChanges();
                    logger.Debug("No. of update plan form =" + i);
                    return i;

                }
                catch (Exception e)
                {
                    logger.Error("update  plan  form  fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }
            }
            return i;
        }

        //新增供應商採購詢價單
        public string addSupplierForm(PLAN_SUP_INQUIRY sf, string[] lstItemId)
        {
            string message = "";
            string sno_key = "PP";
            SerialKeyService snoservice = new SerialKeyService();
            sf.INQUIRY_FORM_ID = snoservice.getSerialKey(sno_key);
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.PLAN_SUP_INQUIRY.AddOrUpdate(sf);
                    i = context.SaveChanges();
                    List<topmeperp.Models.PLAN_SUP_INQUIRY_ITEM> lstItem = new List<PLAN_SUP_INQUIRY_ITEM>();
                    string ItemId = "";
                    for (i = 0; i < lstItemId.Count(); i++)
                    {
                        if (i < lstItemId.Count() - 1)
                        {
                            ItemId = ItemId + "'" + lstItemId[i] + "'" + ",";
                        }
                        else
                        {
                            ItemId = ItemId + "'" + lstItemId[i] + "'";
                        }
                    }

                    string sql = "INSERT INTO PLAN_SUP_INQUIRY_ITEM (INQUIRY_FORM_ID, PLAN_ITEM_ID,"
                        + "TYPE_CODE, SUB_TYPE_CODE,ITEM_DESC,ITEM_UNIT, ITEM_QTY,"
                        + "ITEM_UNIT_PRICE, ITEM_REMARK) "
                        + "SELECT '" + sf.INQUIRY_FORM_ID + "' as INQUIRY_FORM_ID, PLAN_ITEM_ID, TYPE_CODE,"
                        + "SUB_TYPE_CODE, ITEM_DESC, ITEM_UNIT,ITEM_QTY, ITEM_UNIT_PRICE,"
                        + "ITEM_REMARK "
                        + "FROM PLAN_SUP_INQUIRY_ITEM where INQUIRY_ITEM_ID IN (" + ItemId + ")";

                    logger.Info("sql =" + sql);
                    var parameters = new List<SqlParameter>();
                    i = context.Database.ExecuteSqlCommand(sql);

                }
                catch (Exception e)
                {
                    logger.Error("add new plan supplier form id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return sf.INQUIRY_FORM_ID;
        }


        PLAN_SUP_INQUIRY form = null;
        //更新供應商採購詢價單資料
        public int refreshPlanSupplierForm(string formid, PLAN_SUP_INQUIRY sf, List<PLAN_SUP_INQUIRY_ITEM> lstItem)
        {
            logger.Info("Update plan supplier inquiry form id =" + formid);
            form = sf;
            int i = 0;
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.Entry(form).State = EntityState.Modified;
                    i = context.SaveChanges();
                    logger.Debug("Update plan supplier inquiry form =" + i);
                    logger.Info("supplier inquiry form item = " + lstItem.Count);
                    //1.移除詢價單明細
                    deletePlanSupInquiryItem(formid);
                    //2.將item資料寫入 
                    foreach (PLAN_SUP_INQUIRY_ITEM item in lstItem)
                    {
                        PLAN_SUP_INQUIRY_ITEM existItem = null;
                        logger.Debug("form item id=" + item.INQUIRY_ITEM_ID);
                        if (item.INQUIRY_ITEM_ID != 0)
                        {
                            existItem = context.PLAN_SUP_INQUIRY_ITEM.Find(item.INQUIRY_ITEM_ID);
                            if (existItem != null)
                            {
                                existItem.ITEM_QTY = item.ITEM_QTY;
                                existItem.ITEM_UNIT_PRICE = item.ITEM_UNIT_PRICE;
                                existItem.ITEM_REMARK = item.ITEM_REMARK;
                                context.PLAN_SUP_INQUIRY_ITEM.AddOrUpdate(existItem);
                            }
                            else
                            {
                                item.INQUIRY_FORM_ID = sf.INQUIRY_FORM_ID;
                                context.PLAN_SUP_INQUIRY_ITEM.AddOrUpdate(item);
                            }
                        }
                        else
                        {
                            var parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("formid", formid));
                            parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                            string sql = "SELECT * FROM PLAN_SUP_INQUIRY_ITEM WHERE INQUIRY_FORM_ID=@formid AND PLAN_ITEM_ID=@itemid";
                            logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);

                            PLAN_SUP_INQUIRY_ITEM excelItem = null;
                            try
                            {
                                excelItem = context.PLAN_SUP_INQUIRY_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                                existItem = context.PLAN_SUP_INQUIRY_ITEM.Find(excelItem.INQUIRY_ITEM_ID);
                                logger.Debug("find exist item=" + existItem.ITEM_DESC);
                                existItem.ITEM_QTY = item.ITEM_QTY;
                                existItem.ITEM_UNIT_PRICE = item.ITEM_UNIT_PRICE;
                                existItem.ITEM_REMARK = item.ITEM_REMARK;
                                context.PLAN_SUP_INQUIRY_ITEM.AddOrUpdate(existItem);
                            }
                            catch (Exception ex)
                            {
                                //若原來詢價單沒有此項目，將其加入
                                logger.Error(ex.Message + ":" + ex.StackTrace);
                                item.INQUIRY_FORM_ID = formid;
                                context.PLAN_SUP_INQUIRY_ITEM.Add(item);
                            }
                        }
                        j = context.SaveChanges();
                    }

                    logger.Debug("Update plan supplier inquiry form item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new plan supplier form id fail:" + e.Message);
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }
            }
            return i;
        }
        //清空詢價單明細資料
        protected void deletePlanSupInquiryItem(string formid)
        {
            string sqlDel = "DELETE PLAN_SUP_INQUIRY_ITEM WHERE INQUIRY_FORM_ID=@formId";
            logger.Info("sql=" + sqlDel + ",formId=" + formid);
            try
            {
                using (var context = new topmepEntities())
                {
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("formId", formid));
                    context.Database.ExecuteSqlCommand(sqlDel, parameters.ToArray());
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.StackTrace);
            }
        }
        //更新採購詢價單單價
        public int refreshSupplierFormItem(string formid, List<PLAN_SUP_INQUIRY_ITEM> lstItem)
        {
            logger.Info("Update plan supplier inquiry form id =" + formid);
            int j = 0;
            using (var context = new topmepEntities())
            {
                //將item單價寫入 
                foreach (PLAN_SUP_INQUIRY_ITEM item in lstItem)
                {
                    PLAN_SUP_INQUIRY_ITEM existItem = null;
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("formid", formid));
                    parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                    string sql = "SELECT * FROM PLAN_SUP_INQUIRY_ITEM WHERE INQUIRY_FORM_ID=@formid AND PLAN_ITEM_ID=@itemid";
                    logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                    PLAN_SUP_INQUIRY_ITEM excelItem = context.PLAN_SUP_INQUIRY_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                    existItem = context.PLAN_SUP_INQUIRY_ITEM.Find(excelItem.INQUIRY_ITEM_ID);
                    logger.Debug("find exist item=" + existItem.ITEM_DESC);
                    existItem.ITEM_UNIT_PRICE = item.ITEM_UNIT_PRICE;
                    existItem.ITEM_REMARK = item.ITEM_REMARK;
                    context.PLAN_SUP_INQUIRY_ITEM.AddOrUpdate(existItem);
                }
                j = context.SaveChanges();
                logger.Debug("Update plan supplier inquiry form item =" + j);
            }
            return j;
        }
        public int createPlanFormFromSupplier(PLAN_SUP_INQUIRY form, List<PLAN_SUP_INQUIRY_ITEM> items)
        {
            int i = 0;
            //1.建立詢價單價單樣本
            string sno_key = "PP";
            SerialKeyService snoservice = new SerialKeyService();
            form.INQUIRY_FORM_ID = snoservice.getSerialKey(sno_key);
            logger.Info("Plan form from supplier =" + form.ToString());
            using (var context = new topmepEntities())
            {
                context.PLAN_SUP_INQUIRY.Add(form);

                logger.Info("plan form id = " + form.INQUIRY_FORM_ID);
                //if (i > 0) { status = true; };
                foreach (PLAN_SUP_INQUIRY_ITEM item in items)
                {
                    item.INQUIRY_FORM_ID = form.INQUIRY_FORM_ID;
                    context.PLAN_SUP_INQUIRY_ITEM.Add(item);
                }
                i = context.SaveChanges();
            }
            return i;
        }
        public List<string> getSystemMain(string projectid)
        {
            List<string> lst = new List<string>();
            using (var context = new topmepEntities())
            {
                //取得主系統選單
                lst = context.Database.SqlQuery<string>("SELECT DISTINCT SYSTEM_MAIN FROM PLAN_ITEM　WHERE PROJECT_ID=@projectid;", new SqlParameter("projectid", projectid)).ToList();
                logger.Info("Get System Main Count=" + lst.Count);
            }
            return lst;
        }

        //取得材料合約供應商選單
        public List<string> getSupplierForContract(string projectid)
        {
            List<string> lst = new List<string>();
            using (var context = new topmepEntities())
            {
                //取得供應商選單
                lst = context.Database.SqlQuery<string>("SELECT DISTINCT SUPPLIER_ID FROM PLAN_ITEM WHERE PROJECT_ID=@projectid AND SUPPLIER_ID IS NOT NULL ;", new SqlParameter("projectid", projectid)).ToList();
                logger.Info("Get Supplier For Contract Count=" + lst.Count);
            }
            return lst;
        }

        //判斷詢價單是否已被寫入PLAN_ITEM(詢價單已為發包採用)
        public Boolean getSupplierContractByFormId(string formid)
        {
            logger.Info("get boolean of formid in the plan item by formid=" + formid);
            //處理SQL 預先填入ID,設定集合處理參數
            Boolean count = false;
            using (var context = new topmepEntities())
            {
                count = context.Database.SqlQuery<Boolean>("SELECT CAST(COUNT(*) AS BIT) AS BOOLEAN FROM PLAN_ITEM WHERE INQUIRY_FORM_ID =@formid OR MAN_FORM_ID =@formid  ; "
            , new SqlParameter("formid", formid)).FirstOrDefault();
            }

            return count;
        }
        //取得次系統選單
        public List<string> getSystemSub(string projectid)
        {
            List<string> lst = new List<string>();
            using (var context = new topmepEntities())
            {
                //取得主系統選單
                lst = context.Database.SqlQuery<string>("SELECT DISTINCT SYSTEM_SUB FROM PLAN_ITEM WHERE PROJECT_ID=@projectid;", new SqlParameter("projectid", projectid)).ToList();
                //lst = context.TND_PROJECT_ITEM.SqlQuery("SELECT DISTINCT SYSTEM_SUB FROM TND_PROJECT_ITEM　WHERE PROJECT_ID=@projectid;", new SqlParameter("projectid", projectid)).ToList();
                logger.Info("Get System Sub Count=" + lst.Count);
            }
            return lst;
        }

        //取得個別廠商合約資料與金額
        public List<plansummary> getPlanContract(string projectid)
        {
            return getContractForm(projectid, "N");
        }
        //取得個別工資廠商合約資料與金額
        public List<plansummary> getPlanContract4Wage(string projectid)
        {
            return getContractForm(projectid, "Y");
        }

        private static List<plansummary> getContractForm(string projectid, string type)
        {
            List<plansummary> lst = new List<plansummary>();
            using (var context = new topmepEntities())
            {
                string sql4Material = @"SELECT IIF(TYPE='Y','工資','材料') TYPE,VAL.INQUIRY_FORM_ID CONTRACT_ID,
                 VAL.FORM_NAME,VAL.SUPPLIER_ID, 
(SELECT BUDGET_AMOUNT  FROM PLAN_BUDGET 
WHERE PROJECT_ID=@projectid  AND ('('+ TYPE_CODE_1 + ','+TYPE_CODE_2 +')' + BUDGET_NAME)=VAL.FORM_NAME) BUDGET,
                 MATERIAL_COST
                 FROM  (SELECT  * FROM PLAN_ITEM2_SUP_INQUIRY WHERE PROJECT_ID = @projectid ) IDX 
                 INNER JOIN (
                 SELECT F.INQUIRY_FORM_ID,F.FORM_NAME,F.SUPPLIER_ID,SUM(FI.ITEM_QTY * FI.ITEM_UNIT_PRICE) MATERIAL_COST,
                 ISNULL(F.ISWAGE,'N') TYPE
                 FROM PLAN_SUP_INQUIRY F ,PLAN_SUP_INQUIRY_ITEM FI
                 WHERE F.INQUIRY_FORM_ID=FI.INQUIRY_FORM_ID
                 AND F.PROJECT_ID = @projectid 
                 AND ISNULL(SUPPLIER_ID,'') !='' 
                 AND ISNULL(F.ISWAGE,'N')=@type
                 AND ISNULL(F.STATUS,'有效')='有效'
                 GROUP BY F.INQUIRY_FORM_ID,F.FORM_NAME,F.SUPPLIER_ID,ISNULL(F.ISWAGE,'N') ) VAL
                 ON IDX.INQUIRY_FORM_ID=VAL.INQUIRY_FORM_ID";

                lst = context.Database.SqlQuery<plansummary>(sql4Material, new SqlParameter("projectid", projectid), new SqlParameter("type", type)).ToList();
                context.Database.Log = (log) => logger.Debug(log);
            }
            return lst;
        }

        //取得專案廠商合約之金額總計
        public plansummary getPlanContractAmount(string projectid)
        {
            plansummary lst = new plansummary();
            using (var context = new topmepEntities())
            {
                lst = context.Database.SqlQuery<plansummary>("SELECT SUM(REVENUE) TOTAL_REVENUE, SUM(COST) TOTAL_COST, SUM(BUDGET) TOTAL_BUDGET, SUM(PROFIT) TOTAL_PROFIT " +
                    "FROM (select p.FORM_NAME, sum(p.ITEM_QUANTITY * p.ITEM_UNIT_PRICE) REVENUE, " +
                    "sum(map.QTY * tpi.ITEM_UNIT_PRICE * ISNULL(p.BUDGET_RATIO,100) / 100) + sum(map.QTY * w.RATIO * tp.WAGE_MULTIPLIER * ISNULL(p.BUDGET_WAGE_RATIO,100) / 100) BUDGET, " +
                    "(sum(map.QTY * p.ITEM_UNIT_COST) + SUM(map.QTY * ISNULL(p.MAN_PRICE, 0))) COST, (sum(p.ITEM_QUANTITY * p.ITEM_UNIT_PRICE) - sum(map.QTY * p.ITEM_UNIT_COST) - SUM(map.QTY * ISNULL(p.MAN_PRICE, 0))) PROFIT " +
                    "FROM PLAN_ITEM p LEFT JOIN vw_MAP_MATERLIALIST map ON p.PLAN_ITEM_ID = map.PROJECT_ITEM_ID " +
                    "LEFT JOIN TND_PROJECT_ITEM tpi ON p.PLAN_ITEM_ID = tpi.PROJECT_ITEM_ID LEFT JOIN TND_WAGE w ON p.PLAN_ITEM_ID = w.PROJECT_ITEM_ID LEFT JOIN TND_PROJECT tp ON p.PROJECT_ID = tp.PROJECT_ID WHERE p.PROJECT_ID = @projectid " +
                    "GROUP BY p.FORM_NAME)A ; "
                   , new SqlParameter("projectid", projectid)).First();
            }
            return lst;
        }

        //取得採購競標之詢價單資料
        public List<purchasesummary> getPurchaseForm4Offer(string projectid, string formname, string iswage)
        {

            logger.Info("search purchase form by 採購項目 =" + formname);
            List<purchasesummary> lstForm = new List<purchasesummary>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT C.code1 AS FORM_NAME, C.INQUIRY_FORM_ID as INQUIRY_FORM_ID, C.SUPPLIER_ID AS SUPPLIER_ID, D.TOTAL_ROW AS TOTALROWS, D. TOTALPRICE AS TAmount " +
                         "FROM(select p.SUPPLIER_ID, p.INQUIRY_FORM_ID, p.FORM_NAME AS code1, ISNULL(STATUS, '有效') STATUS, ISNULL(ISWAGE, 'N')ISWAGE FROM PLAN_SUP_INQUIRY p LEFT OUTER JOIN PLAN_SUP_INQUIRY_ITEM pi " +
                         "ON p.INQUIRY_FORM_ID = pi.INQUIRY_FORM_ID where p.PROJECT_ID = @projectid AND p.SUPPLIER_ID IS NOT NULL AND ISNULL(STATUS, '有效') <> '註銷' AND ISWAGE <> 'Y' GROUP BY p.FORM_NAME, p.INQUIRY_FORM_ID, " +
                         "p.STATUS, p.SUPPLIER_ID, p.ISWAGE)C LEFT OUTER JOIN (select p.FORM_NAME as type, p.INQUIRY_FORM_ID, count(*) TOTAL_ROW, sum(ITEM_QTY * pi.ITEM_UNIT_PRICE) TOTALPRICE from PLAN_SUP_INQUIRY_ITEM pi " +
                         "LEFT JOIN PLAN_SUP_INQUIRY p ON pi.INQUIRY_FORM_ID = p.INQUIRY_FORM_ID where p.PROJECT_ID = @projectid AND p.SUPPLIER_ID IS NOT NULL GROUP BY p.INQUIRY_FORM_ID, p.FORM_NAME)D " +
                         "ON C.INQUIRY_FORM_ID + C.code1 = D.INQUIRY_FORM_ID + D.type ";
            if (iswage == "Y")
            {
                sql = "SELECT C.code1 AS FORM_NAME, C.INQUIRY_FORM_ID as INQUIRY_FORM_ID, C.SUPPLIER_ID AS SUPPLIER_ID, D.TOTAL_ROW AS TOTALROWS, D.TOTALPRICE AS TAmount " +
                      "FROM(select p.SUPPLIER_ID, p.INQUIRY_FORM_ID, p.FORM_NAME AS code1, ISNULL(STATUS, '有效') STATUS, ISNULL(ISWAGE, 'N')ISWAGE FROM PLAN_SUP_INQUIRY p LEFT OUTER JOIN PLAN_SUP_INQUIRY_ITEM pi " +
                      "ON p.INQUIRY_FORM_ID = pi.INQUIRY_FORM_ID where p.PROJECT_ID =@projectid AND p.SUPPLIER_ID IS NOT NULL AND ISNULL(STATUS, '有效') <> '註銷' AND ISWAGE = 'Y' GROUP BY p.FORM_NAME, p.INQUIRY_FORM_ID, " +
                      "p.STATUS, p.SUPPLIER_ID, p.ISWAGE) C LEFT OUTER JOIN (select p.FORM_NAME as type, p.INQUIRY_FORM_ID, count(*) TOTAL_ROW, sum(ITEM_QTY * pi.ITEM_UNIT_PRICE) TOTALPRICE from PLAN_SUP_INQUIRY_ITEM pi " +
                      "LEFT JOIN PLAN_SUP_INQUIRY p ON pi.INQUIRY_FORM_ID = p.INQUIRY_FORM_ID where p.PROJECT_ID =@projectid AND p.SUPPLIER_ID IS NOT NULL  GROUP BY p.INQUIRY_FORM_ID, p.FORM_NAME)D " +
                      "ON C.INQUIRY_FORM_ID + C.code1 = D.INQUIRY_FORM_ID + D.type ";
            }
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectid));

            //採購項目查詢條件
            if (null != formname && formname != "")
            {
                sql = sql + "WHERE C.code1 =@formname ";
                parameters.Add(new SqlParameter("formname", formname));
            }
            sql = sql + " ORDER BY D.TOTALPRICE ;";
            using (var context = new topmepEntities())
            {
                logger.Debug("get purchase form sql=" + sql);
                lstForm = context.Database.SqlQuery<purchasesummary>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get purchase form count=" + lstForm.Count);
            return lstForm;
        }

        //取得特定專案報價之供應商資料
        public List<COMPARASION_DATA_4PLAN> getComparisonData(string projectid, string typecode1, string typecode2, string systemMain, string systemSub, string formName, string iswage)
        {
            List<COMPARASION_DATA_4PLAN> lst = new List<COMPARASION_DATA_4PLAN>();
            string sql = @"SELECT  pfItem.INQUIRY_FORM_ID AS INQUIRY_FORM_ID, 
                f.SUPPLIER_ID as SUPPLIER_NAME, 
                f.FORM_NAME AS FORM_NAME, 
                ISNULL(f.STATUS,'有效') STATUS,
                ISNULL(SUM(pfitem.ITEM_UNIT_PRICE*pfitem.ITEM_QTY),0) as TAmount, 
                ISNULL(CEILING(SUM(pfitem.ITEM_UNIT_PRICE*pfitem.ITEM_QTY) / SUM(w.RATIO*map.QTY*ISNULL(p.WAGE_MULTIPLIER, -1))),0) as AvgMPrice 
                FROM PLAN_ITEM pItem LEFT OUTER JOIN 
                PLAN_SUP_INQUIRY_ITEM pfItem ON pItem.PLAN_ITEM_ID = pfItem.PLAN_ITEM_ID 
                inner join PLAN_SUP_INQUIRY f on pfItem.INQUIRY_FORM_ID = f.INQUIRY_FORM_ID 
                left outer join TND_WAGE w on pItem.PLAN_ITEM_ID = w.PROJECT_ITEM_ID 
                LEFT outer JOIN vw_MAP_MATERLIALIST map ON pItem.PLAN_ITEM_ID = map.PROJECT_ITEM_ID 
                LEFT JOIN TND_PROJECT p ON pItem.PROJECT_ID = p.PROJECT_ID 
                WHERE pItem.PROJECT_ID = @projectid AND f.SUPPLIER_ID is not null 
                AND ISNULL(f.STATUS,'有效') <> '註銷' AND ISNULL(f.ISWAGE,'N')=@iswage  ";
            var parameters = new List<SqlParameter>();
            //設定專案名編號資料
            parameters.Add(new SqlParameter("projectid", projectid));
            //設定報價單條件，材料或工資
            parameters.Add(new SqlParameter("iswage", iswage));
            //九宮格條件
            if (null != typecode1 && "" != typecode1)
            {
                //sql = sql + " AND pItem.TYPE_CODE_1='" + typecode1 + "'";
                sql = sql + " AND pItem.TYPE_CODE_1=@typecode1";
                parameters.Add(new SqlParameter("typecode1", typecode1));
            }
            //次九宮格條件
            if (null != typecode2 && "" != typecode2)
            {
                //sql = sql + " AND pItem.TYPE_CODE_2='" + typecode2 + "'";
                sql = sql + " AND pItem.TYPE_CODE_2=@typecode2";
                parameters.Add(new SqlParameter("typecode2", typecode2));
            }
            //主系統條件
            if (null != systemMain && "" != systemMain)
            {
                // sql = sql + " AND pItem.SYSTEM_MAIN='" + systemMain + "'";
                sql = sql + " AND pItem.SYSTEM_MAIN=@systemMain";
                parameters.Add(new SqlParameter("systemMain", systemMain));
            }
            //次系統條件
            if (null != systemSub && "" != systemSub)
            {
                //sql = sql + " AND pItem.SYSTEM_SUB='" + systemSub + "'";
                sql = sql + " AND pItem.SYSTEM_SUB=@systemSub";
                parameters.Add(new SqlParameter("systemSub", systemSub));
            }
            //採購名稱條件
            if (null != formName && "" != formName)
            {
                //sql = sql + " AND f.FORM_NAME='" + formName + "'";
                sql = sql + " AND f.FORM_NAME = @formName ";
                parameters.Add(new SqlParameter("formName", formName));
            }
            sql = sql + " GROUP BY pfItem.INQUIRY_FORM_ID ,f.SUPPLIER_ID, f.FORM_NAME, f.STATUS ;";
            logger.Info("comparison data sql=" + sql);
            using (var context = new topmepEntities())
            {
                //取得主系統選單
                lst = context.Database.SqlQuery<COMPARASION_DATA_4PLAN>(sql, parameters.ToArray()).ToList();
                logger.Info("Get ComparisonData Count=" + lst.Count);
            }
            return lst;
        }

        //比價資料
        public DataTable getComparisonDataToPivot(string projectid, string typecode1, string typecode2, string systemMain, string systemSub, string formName, string iswage)
        {
            //採購名稱條件
            if (null != formName && "" != formName)
            {
                string sql = @"SELECT * from (
            select pitem.EXCEL_ROW_ID 行數, 
            pitem.PLAN_ITEM_ID 代號,
            pitem.ITEM_ID 項次,
            pitem.ITEM_DESC 品項名稱,
            pitem.ITEM_UNIT 單位, 
            fitem.ITEM_QTY 數量, 
            (
            SELECT SUPPLIER_ID+'|'+ fitem.INQUIRY_FORM_ID +'|' + FORM_NAME FROM PLAN_SUP_INQUIRY f 
            WHERE f.INQUIRY_FORM_ID = fitem.INQUIRY_FORM_ID 
            ) as SUPPLIER_NAME, 
          pitem.ITEM_UNIT_COST 材料單價, 
          (SELECT FORM_NAME FROM PLAN_SUP_INQUIRY f WHERE f.INQUIRY_FORM_ID = fitem.INQUIRY_FORM_ID ) as FORM_NAME, 
        fitem.ITEM_UNIT_PRICE  
          from PLAN_ITEM pitem 
          left join PLAN_SUP_INQUIRY_ITEM fitem 
           on pitem.PLAN_ITEM_ID = fitem.PLAN_ITEM_ID 
          where pitem.PROJECT_ID = @projectid ";

                if (iswage == "Y")
                {
                    sql = "SELECT * from (select pitem.EXCEL_ROW_ID 行數, pitem.PLAN_ITEM_ID 代號,pitem.ITEM_ID 項次,pitem.ITEM_DESC 品項名稱,pitem.ITEM_UNIT 單位, fitem.ITEM_QTY 數量, " +
                    "(SELECT SUPPLIER_ID+'|'+ fitem.INQUIRY_FORM_ID +'|' + FORM_NAME FROM PLAN_SUP_INQUIRY f WHERE f.INQUIRY_FORM_ID = fitem.INQUIRY_FORM_ID) as SUPPLIER_NAME, " +
                    "pitem.MAN_PRICE 工資單價, " +
                    "(SELECT FORM_NAME FROM PLAN_SUP_INQUIRY f WHERE f.INQUIRY_FORM_ID = fitem.INQUIRY_FORM_ID ) as FORM_NAME, fitem.ITEM_UNIT_PRICE  " +
                    "from PLAN_ITEM pitem " +
                    "left join PLAN_SUP_INQUIRY_ITEM fitem " +
                    " on pitem.PLAN_ITEM_ID = fitem.PLAN_ITEM_ID " +
                    "where pitem.PROJECT_ID = @projectid ";
                }

                var parameters = new Dictionary<string, Object>();
                //設定專案名編號資料
                parameters.Add("projectid", projectid);
                //九宮格條件
                if (null != typecode1 && "" != typecode1)
                {
                    //sql = sql + " AND pItem.TYPE_CODE_1='" + typecode1 + "'";
                    sql = sql + " AND pItem.TYPE_CODE_1=@typecode1";
                    parameters.Add("typecode1", typecode1);
                }
                //次九宮格條件
                if (null != typecode2 && "" != typecode2)
                {
                    //sql = sql + " AND pItem.TYPE_CODE_2='" + typecode2 + "'";
                    sql = sql + " AND pItem.TYPE_CODE_2=@typecode2";
                    parameters.Add("typecode2", typecode2);
                }
                //主系統條件
                if (null != systemMain && "" != systemMain)
                {
                    // sql = sql + " AND pItem.SYSTEM_MAIN='" + systemMain + "'";
                    sql = sql + " AND pItem.SYSTEM_MAIN=@systemMain";
                    parameters.Add("systemMain", systemMain);
                }
                //次系統條件
                if (null != systemSub && "" != systemSub)
                {
                    //sql = sql + " AND pItem.SYSTEM_SUB='" + systemSub + "'";
                    sql = sql + " AND pItem.SYSTEM_SUB=@systemSub";
                    parameters.Add("systemSub", systemSub);
                }

                //取的欄位維度條件
                List<COMPARASION_DATA_4PLAN> lstSuppluerQuo = getComparisonData(projectid, typecode1, typecode2, systemMain, systemSub, formName, iswage);
                if (lstSuppluerQuo.Count == 0)
                {
                    throw new Exception("相關條件沒有任何報價資料!!");
                }
                //設定供應商報價資料，供前端畫面調用
                dirSupplierQuo = new Dictionary<string, COMPARASION_DATA_4PLAN>();
                string dimString = "";
                foreach (var it in lstSuppluerQuo)
                {
                    logger.Debug("Supplier=" + it.SUPPLIER_NAME + "," + it.INQUIRY_FORM_ID + "," + it.FORM_NAME);
                    if (dimString == "")
                    {
                        dimString = "[" + it.SUPPLIER_NAME + "|" + it.INQUIRY_FORM_ID + "|" + it.FORM_NAME + "]";
                    }
                    else
                    {
                        dimString = dimString + ",[" + it.SUPPLIER_NAME + "|" + it.INQUIRY_FORM_ID + "|" + it.FORM_NAME + "]";
                    }
                    //設定供應商報價資料，供前端畫面調用
                    dirSupplierQuo.Add(it.INQUIRY_FORM_ID, it);
                }

                logger.Debug("dimString=" + dimString);
                //sql = sql + " AND FORM_NAME ='" + formName + "'";
                sql = sql + ") souce pivot(MIN(ITEM_UNIT_PRICE) FOR SUPPLIER_NAME IN(" + dimString + ")) as pvt WHERE FORM_NAME =@formName ORDER BY 行數; ";
                parameters.Add("formName", formName);
                logger.Info("comparison data sql=" + sql);
                DataSet ds = ExecuteStoreQuery(sql, CommandType.Text, parameters);
                //Pivot pvt = new Pivot(ds.Tables[0]);
                return ds.Tables[0];
            }
            else
            {
                string sql = "SELECT * from (select pitem.EXCEL_ROW_ID 行數, pitem.PLAN_ITEM_ID 代號,pitem.ITEM_ID 項次,pitem.ITEM_DESC 品項名稱,pitem.ITEM_UNIT 單位,fitem.ITEM_QTY 數量, " +
                "(SELECT SUPPLIER_ID+'|'+ fitem.INQUIRY_FORM_ID +'|' + FORM_NAME FROM PLAN_SUP_INQUIRY f WHERE f.INQUIRY_FORM_ID = fitem.INQUIRY_FORM_ID ) as SUPPLIER_NAME, " +
                "pitem.ITEM_UNIT_COST 材料單價, fitem.ITEM_UNIT_PRICE " +
                "from PLAN_ITEM pitem " +
                "left join PLAN_SUP_INQUIRY_ITEM fitem " +
                " on pitem.PLAN_ITEM_ID = fitem.PLAN_ITEM_ID " +
                "where pitem.PROJECT_ID = @projectid ";

                if (iswage == "Y")
                {
                    sql = "SELECT * from (select pitem.EXCEL_ROW_ID 行數, pitem.PLAN_ITEM_ID 代號,pitem.ITEM_ID 項次,pitem.ITEM_DESC 品項名稱,pitem.ITEM_UNIT 單位,fitem.ITEM_QTY 數量, " +
                    "(SELECT SUPPLIER_ID+'|'+ fitem.INQUIRY_FORM_ID +'|' + FORM_NAME FROM PLAN_SUP_INQUIRY f WHERE f.INQUIRY_FORM_ID = fitem.INQUIRY_FORM_ID) as SUPPLIER_NAME, " +
                    "pitem.MAN_PRICE 工資單價, fitem.ITEM_UNIT_PRICE " +
                    "from PLAN_ITEM pitem " +
                    "left join PLAN_SUP_INQUIRY_ITEM fitem " +
                    " on pitem.PLAN_ITEM_ID = fitem.PLAN_ITEM_ID " +
                    "where pitem.PROJECT_ID = @projectid ";
                }
                var parameters = new Dictionary<string, Object>();
                //設定專案名編號資料
                parameters.Add("projectid", projectid);
                //九宮格條件
                if (null != typecode1 && "" != typecode1)
                {
                    //sql = sql + " AND pItem.TYPE_CODE_1='" + typecode1 + "'";
                    sql = sql + " AND pItem.TYPE_CODE_1=@typecode1";
                    parameters.Add("typecode1", typecode1);
                }
                //次九宮格條件
                if (null != typecode2 && "" != typecode2)
                {
                    //sql = sql + " AND pItem.TYPE_CODE_2='" + typecode2 + "'";
                    sql = sql + " AND pItem.TYPE_CODE_2=@typecode2";
                    parameters.Add("typecode2", typecode2);
                }
                //主系統條件
                if (null != systemMain && "" != systemMain)
                {
                    // sql = sql + " AND pItem.SYSTEM_MAIN='" + systemMain + "'";
                    sql = sql + " AND pItem.SYSTEM_MAIN=@systemMain";
                    parameters.Add("systemMain", systemMain);
                }
                //次系統條件
                if (null != systemSub && "" != systemSub)
                {
                    //sql = sql + " AND pItem.SYSTEM_SUB='" + systemSub + "'";
                    sql = sql + " AND pItem.SYSTEM_SUB=@systemSub";
                    parameters.Add("systemSub", systemSub);
                }

                //取的欄位維度條件
                List<COMPARASION_DATA_4PLAN> lstSuppluerQuo = getComparisonData(projectid, typecode1, typecode2, systemMain, systemSub, formName, iswage);
                if (lstSuppluerQuo.Count == 0)
                {
                    throw new Exception("相關條件沒有任何報價資料!!");
                }
                //設定供應商報價資料，供前端畫面調用
                dirSupplierQuo = new Dictionary<string, COMPARASION_DATA_4PLAN>();
                string dimString = "";
                foreach (var it in lstSuppluerQuo)
                {
                    logger.Debug("Supplier=" + it.SUPPLIER_NAME + "," + it.INQUIRY_FORM_ID + "," + it.FORM_NAME);
                    if (dimString == "")
                    {
                        dimString = "[" + it.SUPPLIER_NAME + "|" + it.INQUIRY_FORM_ID + "|" + it.FORM_NAME + "]";
                    }
                    else
                    {
                        dimString = dimString + ",[" + it.SUPPLIER_NAME + "|" + it.INQUIRY_FORM_ID + "|" + it.FORM_NAME + "]";
                    }
                    //設定供應商報價資料，供前端畫面調用
                    dirSupplierQuo.Add(it.INQUIRY_FORM_ID, it);
                }

                logger.Debug("dimString=" + dimString);
                //sql = sql + " AND FORM_NAME ='" + formName + "'";
                sql = sql + ") souce pivot(MIN(ITEM_UNIT_PRICE) FOR SUPPLIER_NAME IN(" + dimString + ")) as pvt ORDER BY 行數; ";
                logger.Info("comparison data sql=" + sql);
                DataSet ds = ExecuteStoreQuery(sql, CommandType.Text, parameters);
                //Pivot pvt = new Pivot(ds.Tables[0]);
                return ds.Tables[0];
            }
        }

        public int addSuplplierFormFromQuote(string formid)
        {
            int i = 0;
            logger.Info("add plan supplier form from Quote by form id" + formid);
            string sql = "UPDATE  PLAN_SUP_INQUIRY SET COUNTER_OFFER = 'Y', MODIFY_DATE = getdate()  WHERE INQUIRY_FORM_ID=@formid ";
            logger.Debug("add form from Quote sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formid", formid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return i;
        }
        public int removeSuplplierFormFromQuote(string formid)
        {
            int i = 0;
            logger.Info("Remove plan supplier form from Quote by form id" + formid);
            string sql = "UPDATE  PLAN_SUP_INQUIRY SET STATUS = '註銷', MODIFY_DATE = getdate() WHERE INQUIRY_FORM_ID=@formid ";
            logger.Debug("remove form from Quote sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formid", formid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return i;
        }

        public int filterSuplplierFormFromQuote(string formid)
        {
            int i = 0;
            logger.Info("Filter plan supplier form from Quote by form id" + formid);
            string sql = "UPDATE  PLAN_SUP_INQUIRY SET COUNTER_OFFER = 'M', MODIFY_DATE = getdate() WHERE INQUIRY_FORM_ID=@formid ";
            logger.Debug("remove form from Quote sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formid", formid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return i;
        }
        //由報價單資料更新標單資料
        public int updateCostFromQuote(string planItemid, decimal price, string iswage)
        {
            int i = 0;
            logger.Info("Update Cost:plan item id=" + planItemid + ",price=" + price);
            db = new topmepEntities();
            string sql = "UPDATE PLAN_ITEM SET ITEM_UNIT_COST =@price WHERE PLAN_ITEM_ID=@pitemid ";
            //將工資報價單更新工資報價欄位
            if (iswage == "Y")
            {
                sql = "UPDATE PLAN_ITEM SET MAN_PRICE=@price WHERE PLAN_ITEM_ID=@pitemid ";
            }
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("price", price));
            parameters.Add(new SqlParameter("pitemid", planItemid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            db = null;
            logger.Info("Update Cost:" + i);
            return i;
        }
        //將發包廠商之詢價單價格寫入PLAN_ITEM且寫入後不得再覆寫
        public int batchUpdateCostFromQuote(string formid, string iswage)
        {
            int i = 0;
            logger.Info("Copy cost from Quote to Tnd by form id" + formid);
            string sql = "UPDATE  PLAN_ITEM SET item_unit_cost = i.ITEM_UNIT_PRICE, supplier_id = i.SUPPLIER_ID, form_name = i.FORM_NAME, inquiry_form_id = i.INQUIRY_FORM_ID " +
                "FROM(select i.plan_item_id, fi.ITEM_UNIT_PRICE, fi.INQUIRY_FORM_ID, pf.SUPPLIER_ID, pf.FORM_NAME from PLAN_ITEM i " +
                ", PLAN_SUP_INQUIRY_ITEM fi, PLAN_SUP_INQUIRY pf " +
               "where i.PLAN_ITEM_ID = fi.PLAN_ITEM_ID and fi.INQUIRY_FORM_ID = pf.INQUIRY_FORM_ID and fi.INQUIRY_FORM_ID = @formid) i " +
                "WHERE  i.plan_item_id = PLAN_ITEM.PLAN_ITEM_ID  ";

            //將工資報價單更新工資報價欄位
            if (iswage == "Y")
            {
                sql = "UPDATE  PLAN_ITEM SET man_price = i.ITEM_UNIT_PRICE, man_supplier_id = i.SUPPLIER_ID, man_form_name = i.FORM_NAME, man_form_id = i.INQUIRY_FORM_ID "
                + "FROM (select i.plan_item_id, fi.ITEM_UNIT_PRICE, fi.INQUIRY_FORM_ID, pf.SUPPLIER_ID, pf.FORM_NAME from PLAN_ITEM i "
                + ", PLAN_SUP_INQUIRY_ITEM fi, PLAN_SUP_INQUIRY pf "
                + "where i.PLAN_ITEM_ID = fi.PLAN_ITEM_ID and fi.INQUIRY_FORM_ID = pf.INQUIRY_FORM_ID and fi.INQUIRY_FORM_ID = @formid) i "
                + "WHERE  i.plan_item_id = PLAN_ITEM.PLAN_ITEM_ID  ";
            }
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formid", formid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return i;
        }
        /// <summary>
        /// 增加供應商發包記錄
        /// </summary>
        /// <param name="projectid"></param>
        /// <param name="formid"></param>
        /// <param name="u"></param>
        /// <returns></returns>
        public int addContract4SupplierRecord(string projectid, string formid, SYS_USER u)
        {
            int i = 0;
            PLAN_ITEM2_SUP_INQUIRY p2s = new PLAN_ITEM2_SUP_INQUIRY();
            p2s.PROJECT_ID = projectid;
            p2s.INQUIRY_FORM_ID = formid;
            p2s.CREATE_ID = u.CREATE_ID;
            p2s.CREATE_DATE = DateTime.Now;
            try
            {
                using (var context = new topmepEntities())
                {
                    context.PLAN_ITEM2_SUP_INQUIRY.Add(p2s);
                    i = context.SaveChanges();
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message + ":" + ex.StackTrace);
            }
            return i;
        }
        public int addContractId(string projectid)
        {
            int i = 0;
            //將材料合約編號寫入PLAN PAYMENT TERMS
            logger.Info("copy contract id into plan payment terms, project id =" + projectid);
            using (var context = new topmepEntities())
            {
                List<topmeperp.Models.PLAN_PAYMENT_TERMS> lstItem = new List<PLAN_PAYMENT_TERMS>();
                string sql = "INSERT INTO PLAN_PAYMENT_TERMS (CONTRACT_ID, PROJECT_ID, TYPE) " +
                       "SELECT distinct A.INQUIRY_FORM_ID AS contractid, '" + projectid + "', 'S' FROM " +
                       "(SELECT pi.PROJECT_ID, pi.SUPPLIER_ID AS SUPPLIER_NAME, pi.INQUIRY_FORM_ID, s.SUPPLIER_ID FROM PLAN_ITEM pi LEFT JOIN TND_SUPPLIER s ON " +
                       "pi.SUPPLIER_ID = s.COMPANY_NAME WHERE pi.SUPPLIER_ID IS NOT NULL)A WHERE A.PROJECT_ID = '" + projectid + "' " +
                       "AND A.INQUIRY_FORM_ID NOT IN(SELECT ppt.CONTRACT_ID FROM PLAN_PAYMENT_TERMS ppt) ";

                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return i;
            }
        }
        //將工資合約編號寫入PLAN PAYMENT TERMS
        public int addContractIdForWage(string projectid)
        {
            int i = 0;
            logger.Info("copy contract id from wage into plan payment terms, project id =" + projectid);
            using (var context = new topmepEntities())
            {
                List<topmeperp.Models.PLAN_PAYMENT_TERMS> lstItem = new List<PLAN_PAYMENT_TERMS>();
                string sql = "INSERT INTO PLAN_PAYMENT_TERMS (CONTRACT_ID, PROJECT_ID, TYPE) " +
                   "SELECT distinct A.MAN_FORM_ID AS contractid, '" + projectid + "', 'S' FROM " +
                   "(SELECT pi.PROJECT_ID, pi.MAN_SUPPLIER_ID AS SUPPLIER_NAME, pi.MAN_FORM_ID, s.SUPPLIER_ID FROM PLAN_ITEM pi LEFT JOIN TND_SUPPLIER s ON " +
                   "pi.MAN_SUPPLIER_ID = s.COMPANY_NAME WHERE pi.MAN_SUPPLIER_ID IS NOT NULL)A WHERE A.PROJECT_ID = '" + projectid + "' " +
                   "AND A.MAN_FORM_ID NOT IN (SELECT ppt.CONTRACT_ID FROM PLAN_PAYMENT_TERMS ppt) ";

                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return i;
            }
        }

        public List<PLAN_ITEM> getContractItemsByContractName(string contractid)
        {
            List<PLAN_ITEM> lst = new List<PLAN_ITEM>();
            using (var context = new topmepEntities())
            {
                lst = context.PLAN_ITEM.SqlQuery("SELECT * FROM PLAN_ITEM WHERE INQUIRY_FORM_ID OR MAN_FORM_ID = @contractid ;"
                    , new SqlParameter("contractid", contractid)).ToList();
            }
            logger.Info("get plan supplier contract items count:" + lst.Count);
            return lst;
        }

        public PaymentTermsFunction getPaymentTerm(string contractid, string estid)
        {
            logger.Debug("get payment terms by contractid=" + contractid + ",estid=" + estid);
            PaymentTermsFunction payment = null;
            using (var context = new topmepEntities())
            {
                //條件篩選
                string sql = @"SELECT ppt.*, IIF(ef.advanceAmt < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT*ppt.PAYMENT_ADVANCE_CASH_RATIO/100,ef.PAID_AMOUNT*USANCE_ADVANCE_CASH_RATIO/100) AS advanceCash, 
                         IIF(ef.advanceAmt < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_ADVANCE_1_RATIO / 100, ef.PAID_AMOUNT * USANCE_ADVANCE_1_RATIO / 100) AS advanceAmt1, 
                         IIF(ef.advanceAmt < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_ADVANCE_2_RATIO / 100, ef.PAID_AMOUNT * USANCE_ADVANCE_2_RATIO / 100) AS advanceAmt2, 
                         IIF(ef.RETENTION_PAYMENT < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_RETENTION_CASH_RATIO / 100, ef.PAID_AMOUNT * USANCE_RETENTION_CASH_RATIO / 100) AS retentionCash, 
                         IIF(ef.RETENTION_PAYMENT < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_RETENTION_1_RATIO / 100, ef.PAID_AMOUNT * USANCE_RETENTION_1_RATIO / 100) AS retentionAmt1, 
                         IIF(ef.RETENTION_PAYMENT < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_RETENTION_2_RATIO / 100, ef.PAID_AMOUNT * USANCE_RETENTION_2_RATIO / 100) AS retentionAmt2,
                         IIF(ef.PAYMENT_TRANSFER > 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * PAYMENT_ESTIMATED_CASH_RATIO / 100, ef.PAID_AMOUNT * USANCE_GOODS_CASH_RATIO / 100) AS estCash, 
                         IIF(ef.PAYMENT_TRANSFER > 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * PAYMENT_ESTIMATED_1_RATIO / 100, ef.PAID_AMOUNT * USANCE_GOODS_1_RATIO / 100) AS estAmt1, 
                         IIF(ef.PAYMENT_TRANSFER > 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * PAYMENT_ESTIMATED_2_RATIO / 100, ef.PAID_AMOUNT * USANCE_GOODS_2_RATIO / 100) AS estAmt2,
                         IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3) AS dateBase1, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null) AS dateBase2,
                         IIF(DAY(ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3),IIF(IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null) is null,IIF(MONTH(ef.CREATE_DATE) > 11, CONVERT(varchar, YEAR(ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)),
                         CONVERT(varchar, YEAR(ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))),IIF(DAY(ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null), IIF(MONTH(ef.CREATE_DATE) > 11, CONVERT(varchar, YEAR(ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), 
                         CONVERT(varchar, YEAR(ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), CONVERT(varchar, YEAR(ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)))), 
                         CONVERT(varchar, YEAR(ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, DATE_3))) AS paidDateCash, 
                         IIF(DAY(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3),IIF(IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null) is null,IIF(MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) > 11, 
                         CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), 
                         CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), 
                         IIF(DAY(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null), IIF(MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) > 11, 
                         CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), 
                         CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), 
                         CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)))), 
                         CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, DATE_3))) AS paidDate1, 
                         IIF(DAY(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3),IIF(IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null) is null,IIF(MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) > 11, 
                         CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), 
                         CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), 
                         IIF(DAY(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null), IIF(MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) > 11, 
                         CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), 
                         CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), 
                         CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)))), 
                         CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, DATE_3))) AS paidDate2 
                         FROM PLAN_PAYMENT_TERMS ppt LEFT JOIN(SELECT ef.CONTRACT_ID, pop.advanceAmt, ef.RETENTION_PAYMENT, ef.PAID_AMOUNT, ef.PAYMENT_TRANSFER, ef.CREATE_DATE FROM PLAN_ESTIMATION_FORM ef LEFT JOIN(SELECT EST_FORM_ID, SUM(AMOUNT) AS advanceAmt FROM PLAN_OTHER_PAYMENT WHERE TYPE IN('A', 'B', 'C') GROUP BY EST_FORM_ID HAVING EST_FORM_ID =@estid)pop 
                         ON ef.EST_FORM_ID = pop.EST_FORM_ID WHERE ef.EST_FORM_ID =@estid)ef ON ppt.CONTRACT_ID = ef.CONTRACT_ID WHERE ppt.CONTRACT_ID =@contractid ";
                logger.Debug("sql=" + sql);
                payment = context.Database.SqlQuery<PaymentTermsFunction>(sql, new SqlParameter("contractid", contractid), new SqlParameter("estid", estid)).FirstOrDefault();
            }
            return payment;
        }

        public List<PlanItem4Map> getPendingItems(string projectid)
        {
            List<PlanItem4Map> lst = new List<PlanItem4Map>();
            using (var context = new topmepEntities())
            {
                //取得材料中有列入分項項目之標單品項但卻未被發包的品項
                lst = context.Database.SqlQuery<PlanItem4Map>("SELECT main.*, contract.INQUIRY_FORM_ID AS formId FROM (SELECT pi.*,fitem.FORM_NAME AS formName FROM PLAN_ITEM pi LEFT JOIN (SELECT it.PLAN_ITEM_ID, psi.FORM_NAME, ISNULL(psi.ISWAGE, 'N') AS isWage FROM PLAN_SUP_INQUIRY_ITEM it LEFT JOIN PLAN_SUP_INQUIRY psi " +
                    "ON it.INQUIRY_FORM_ID = psi.INQUIRY_FORM_ID WHERE psi.PROJECT_ID = @projectid AND psi.SUPPLIER_ID IS NULL AND ISNULL(STATUS, '有效') = '有效' AND psi.ISWAGE <> 'Y') fitem ON pi.PLAN_ITEM_ID = fitem.PLAN_ITEM_ID " +
                    "WHERE PROJECT_ID = @projectid AND pi.FORM_NAME IS NULL AND fitem.PLAN_ITEM_ID IS NOT NULL AND fitem.FORM_NAME + isWage NOT IN (SELECT FORM_NAME + ISNULL(ISWAGE, 'N') AS tag FROM PLAN_SUP_INQUIRY WHERE SUPPLIER_ID IS NULL AND PROJECT_ID = @projectid AND ISNULL(STATUS, '有效') = '有效' AND ISNULL(ISWAGE, 'N') = 'Y' " +
                    "AND FORM_NAME NOT IN (SELECT p.FORM_NAME AS CODE FROM PLAN_ITEM p WHERE p.PROJECT_ID = @projectid AND p.FORM_NAME IS NOT NULL GROUP BY p.FORM_NAME)) " +
                    "OR PROJECT_ID = @projectid AND pi.FORM_NAME = '' AND fitem.PLAN_ITEM_ID IS NOT NULL AND fitem.FORM_NAME + isWage NOT IN (SELECT FORM_NAME + ISNULL(ISWAGE, 'N') AS tag FROM PLAN_SUP_INQUIRY WHERE SUPPLIER_ID IS NULL AND PROJECT_ID = @projectid AND ISNULL(STATUS, '有效') = '有效' AND ISNULL(ISWAGE, 'N') = 'Y' " +
                    "AND FORM_NAME NOT IN (SELECT p.FORM_NAME AS CODE FROM PLAN_ITEM p WHERE p.PROJECT_ID = @projectid AND p.FORM_NAME IS NOT NULL GROUP BY p.FORM_NAME)))main LEFT JOIN (SELECT pi.FORM_NAME, pi.INQUIRY_FORM_ID FROM PLAN_ITEM pi GROUP BY pi.FORM_NAME, pi.INQUIRY_FORM_ID HAVING pi.FORM_NAME IS NOT NULL)contract " +
                    "ON main.formName = contract.FORM_NAME ; ", new SqlParameter("projectid", projectid)).ToList();
            }
            logger.Info("get plan pending items count:" + lst.Count);
            return lst;
        }

        public List<PlanItem4Map> getPendingItems4Wage(string projectid)
        {
            List<PlanItem4Map> lst = new List<PlanItem4Map>();
            using (var context = new topmepEntities())
            {
                //取得工資中有列入分項項目之標單品項但卻未被發包的品項
                lst = context.Database.SqlQuery<PlanItem4Map>("SELECT main.*, contract.MAN_FORM_ID AS formId FROM (SELECT pi.*, fitem.FORM_NAME AS formName FROM PLAN_ITEM pi LEFT JOIN (SELECT it.PLAN_ITEM_ID, psi.FORM_NAME, ISNULL(psi.ISWAGE, 'N') AS isWage FROM PLAN_SUP_INQUIRY_ITEM it LEFT JOIN PLAN_SUP_INQUIRY psi " +
                    "ON it.INQUIRY_FORM_ID = psi.INQUIRY_FORM_ID WHERE psi.PROJECT_ID = @projectid AND psi.SUPPLIER_ID IS NULL AND ISNULL(STATUS, '有效') = '有效' AND psi.ISWAGE = 'Y') fitem ON pi.PLAN_ITEM_ID = fitem.PLAN_ITEM_ID " +
                    "WHERE PROJECT_ID = @projectid AND pi.MAN_FORM_NAME IS NULL AND fitem.PLAN_ITEM_ID IS NOT NULL AND fitem.FORM_NAME + isWage NOT IN (SELECT FORM_NAME + ISNULL(ISWAGE, 'N') AS tag FROM PLAN_SUP_INQUIRY WHERE SUPPLIER_ID IS NULL AND PROJECT_ID = @projectid AND ISNULL(STATUS, '有效') = '有效' AND ISNULL(ISWAGE, 'N') = 'Y' " +
                    "AND FORM_NAME NOT IN (SELECT p.MAN_FORM_NAME AS CODE FROM PLAN_ITEM p WHERE p.PROJECT_ID = @projectid AND p.MAN_FORM_NAME IS NOT NULL GROUP BY p.MAN_FORM_NAME)) " +
                    "OR PROJECT_ID = @projectid AND pi.MAN_FORM_NAME = '' AND fitem.PLAN_ITEM_ID IS NOT NULL AND fitem.FORM_NAME + isWage NOT IN (SELECT FORM_NAME + ISNULL(ISWAGE, 'N') AS tag FROM PLAN_SUP_INQUIRY WHERE SUPPLIER_ID IS NULL AND PROJECT_ID = @projectid AND ISNULL(STATUS, '有效') = '有效' AND ISNULL(ISWAGE, 'N') = 'Y' " +
                    "AND FORM_NAME NOT IN (SELECT p.MAN_FORM_NAME AS CODE FROM PLAN_ITEM p WHERE p.PROJECT_ID = @projectid AND p.MAN_FORM_NAME IS NOT NULL GROUP BY p.MAN_FORM_NAME)))main LEFT JOIN (SELECT pi.MAN_FORM_NAME, pi.MAN_FORM_ID FROM PLAN_ITEM pi GROUP BY pi.MAN_FORM_NAME, pi.MAN_FORM_ID HAVING pi.MAN_FORM_NAME IS NOT NULL)contract " +
                    "ON main.formName = contract.MAN_FORM_NAME ; ", new SqlParameter("projectid", projectid)).ToList();
            }
            logger.Info("get plan pending items count:" + lst.Count);
            return lst;
        }

        public int updatePaymentTerms(PLAN_PAYMENT_TERMS item)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.PLAN_PAYMENT_TERMS.AddOrUpdate(item);
                    i = context.SaveChanges();
                }
                catch (Exception e)
                {
                    logger.Error("updatePaymentTerms fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }
        public int changePlanFormStatus(string formid, string status)
        {
            int i = 0;
            logger.Info("Update plan sup inquiry form status formid=" + formid + ",status=" + status);
            db = new topmepEntities();
            string sql = "UPDATE PLAN_SUP_INQUIRY SET STATUS=@status WHERE INQUIRY_FORM_ID=@formid ";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("status", status));
            parameters.Add(new SqlParameter("formid", formid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            db = null;
            logger.Debug("Update plan sup inquiry form status  :" + i);
            return i;
        }
        //更新得標標單品項之採購前置天數
        public int updateLeadTime(string projectid, List<PLAN_ITEM> lstItem)
        {
            //1.新增前置天數資料
            int i = 0;
            logger.Info("update lead time = " + lstItem.Count);
            //2.將預算資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_ITEM item in lstItem)
                {
                    PLAN_ITEM existItem = null;
                    logger.Debug("plan item id=" + item.PLAN_ITEM_ID);
                    if (item.PLAN_ITEM_ID != null && item.PLAN_ITEM_ID != "")
                    {
                        existItem = context.PLAN_ITEM.Find(item.PLAN_ITEM_ID);
                    }
                    else
                    {
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("projectid", projectid));
                        parameters.Add(new SqlParameter("planitemid", item.PLAN_ITEM_ID));
                        string sql = "SELECT * FROM PLAN_ITEM WHERE PROJECT_ID = @projectid and PLAN_ITEM_ID = @planitemid ";
                        logger.Info(sql + " ;" + item.PROJECT_ID + item.PLAN_ITEM_ID);
                        PLAN_ITEM excelItem = context.PLAN_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_ITEM.Find(excelItem.PLAN_ITEM_ID);

                    }
                    logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                    existItem.LEAD_TIME = item.LEAD_TIME;
                    existItem.MODIFY_USER_ID = item.MODIFY_USER_ID;
                    existItem.MODIFY_DATE = DateTime.Now;
                    context.PLAN_ITEM.AddOrUpdate(existItem);
                }
                i = context.SaveChanges();
            }
            logger.Info("update lead time count =" + i);
            return i;
        }

        #region 物料管理之進銷存
        public PLAN_PURCHASE_REQUISITION formPR = null;
        public List<PurchaseRequisition> PRItem = null;
        public List<PurchaseRequisition> DOItem = null;

        public List<PurchaseRequisition> getPurchaseItemByMap(string projectid, List<string> lstItemId)
        {
            //取得任務採購內容
            logger.Info("get plan item by map ");
            List<PurchaseRequisition> lstItem = new List<PurchaseRequisition>();
            using (var context = new topmepEntities())
            {
                string ItemId = "";
                for (i = 0; i < lstItemId.Count(); i++)
                {
                    if (i < lstItemId.Count() - 1)
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'" + ",";
                    }
                    else
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'";
                    }
                }

                string sql = @"SELECT pi.* , map.QTY AS MAP_QTY, B.CUMULATIVE_QTY, C.ALL_RECEIPT_QTY- D.DELIVERY_QTY AS INVENTORY_QTY 
FROM PLAN_ITEM pi  LEFT OUTER JOIN  vw_MAP_MATERLIALIST map ON pi.PLAN_ITEM_ID = map.PROJECT_ITEM_ID 
LEFT JOIN PLAN_COSTCHANGE_ITEM pci ON pi.PLAN_ITEM_ID = pci.PLAN_ITEM_ID 
LEFT JOIN (SELECT pri.PLAN_ITEM_ID, SUM(pri.ORDER_QTY) AS CUMULATIVE_QTY 
                   FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pri.PR_ID = pr.PR_ID WHERE pr.PROJECT_ID = @projectid AND 
                   pri.PR_ID LIKE 'PPO%' GROUP BY pri.PLAN_ITEM_ID )B ON pi.PLAN_ITEM_ID = B.PLAN_ITEM_ID 
                   LEFT JOIN(SELECT pri.PLAN_ITEM_ID, SUM(pri.RECEIPT_QTY) AS ALL_RECEIPT_QTY FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_PURCHASE_REQUISITION pr 
                   ON pri.PR_ID = pr.PR_ID WHERE pr.PROJECT_ID = @projectid AND pri.PR_ID LIKE 'RP%' GROUP BY pri.PLAN_ITEM_ID)C 
                   ON pi.PLAN_ITEM_ID = C.PLAN_ITEM_ID LEFT JOIN (SELECT pid.PLAN_ITEM_ID, SUM(pid.DELIVERY_QTY) AS DELIVERY_QTY FROM PLAN_ITEM_DELIVERY pid 
                   LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pid.DELIVERY_ORDER_ID = pr.PR_ID WHERE pr.PROJECT_ID = @projectid 
                   GROUP BY pid.PLAN_ITEM_ID)D ON pi.PLAN_ITEM_ID = D.PLAN_ITEM_ID WHERE pi.PROJECT_ID = @projectid AND pi.PLAN_ITEM_ID IN ($ItemList)";
                sql = sql.Replace("$ItemList", ItemId);
                logger.Info("sql = " + sql);
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", projectid));
                lstItem = context.Database.SqlQuery<PurchaseRequisition>(sql, parameters.ToArray()).ToList();
                logger.Info("Get task material Info Record Count=" + lstItem.Count);
            }
            return lstItem;
        }

        // 寫入任務採購內容
        public string newPR(string projectid, PLAN_PURCHASE_REQUISITION form, string[] lstItemId)
        {
            //1.建立申購單
            logger.Info("create new purchase requisition ");
            string sno_key = "PR";
            SerialKeyService snoservice = new SerialKeyService();
            form.PR_ID = snoservice.getSerialKey(sno_key);
            logger.Info("new purchase requisition =" + form.ToString());
            using (var context = new topmepEntities())
            {
                context.PLAN_PURCHASE_REQUISITION.Add(form);
                int i = context.SaveChanges();
                logger.Debug("Add Purchase Requisition=" + i);
                logger.Info("plan purchase requisition id = " + form.PR_ID);
                //if (i > 0) { status = true; };
                List<topmeperp.Models.PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
                string ItemId = "";
                for (i = 0; i < lstItemId.Count(); i++)
                {
                    if (i < lstItemId.Count() - 1)
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'" + ",";
                    }
                    else
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'";
                    }
                }

                string sql = "INSERT INTO PLAN_PURCHASE_REQUISITION_ITEM (PR_ID, PLAN_ITEM_ID) "
                + "SELECT '" + form.PR_ID + "' as PR_ID, A.PLAN_ITEM_ID as PLAN_ITEM_ID  "
                + "FROM (SELECT pi.PLAN_ITEM_ID FROM PLAN_ITEM pi WHERE pi.PLAN_ITEM_ID IN (" + ItemId + "))A ";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return form.PR_ID;
            }
        }
        public void delPR(string prid)
        {
            string sql = @"DELETE PLAN_PURCHASE_REQUISITION_ITEM WHERE PR_ID=@prid;
                          DELETE PLAN_PURCHASE_REQUISITION WHERE PR_ID=@prid";
            logger.Info("sql= " + sql + ",prid=" + prid);
            using (var context = new topmepEntities())
            {
                int i = context.Database.ExecuteSqlCommand(sql, new SqlParameter("prid", prid));
                logger.Info("count= " + i);
            }
        }
        //更新申購數量
        public int refreshPR(string formid, PLAN_PURCHASE_REQUISITION form, List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem)
        {
            logger.Info("Update plan purchase requisition id =" + formid);
            int i = 0;
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.Entry(form).State = EntityState.Modified;
                    i = context.SaveChanges();
                    logger.Debug("Update plan purchase requisition =" + i);
                    logger.Info("purchase requisition item = " + lstItem.Count);
                    //2.將item資料寫入 
                    foreach (PLAN_PURCHASE_REQUISITION_ITEM item in lstItem)
                    {
                        PLAN_PURCHASE_REQUISITION_ITEM existItem = null;
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("formid", formid));
                        parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                        string sql = "SELECT * FROM PLAN_PURCHASE_REQUISITION_ITEM WHERE PR_ID=@formid AND PLAN_ITEM_ID=@itemid";
                        logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                        PLAN_PURCHASE_REQUISITION_ITEM excelItem = context.PLAN_PURCHASE_REQUISITION_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(excelItem.PR_ITEM_ID);
                        logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                        existItem.NEED_QTY = item.NEED_QTY;
                        existItem.NEED_DATE = item.NEED_DATE;
                        existItem.REMARK = item.REMARK;
                        context.PLAN_PURCHASE_REQUISITION_ITEM.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update purchase requisition item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new purchase requisition id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }
        /// <summary>
        /// 讀取申購單、採購單、驗收單、領料單(?) 等資料
        /// </summary>
        /// <returns></returns>
        public List<PRFunction> getPRFunction(string type, string projectid, string status, string yymm, string prid, string keyword)
        {
            // Type =PR : 申購單
            // Type =PPO : 採購單
            // Type= RP : 驗收單
            List<PRFunction> lstForm = null;
            string sql = @"SELECT 
                    ROW_NUMBER() OVER(ORDER BY P.PR_ID desc, P.STATUS,P.PARENT_PR_ID desc) AS NO ,
                    P.PROJECT_ID, 
                    CONVERT(char(10), P.CREATE_DATE, 111) AS CREATE_DATE, 
                    P.PR_ID, 
                    P.STATUS, '' as TASK_NAME, 
                    ISNULL(P.REMARK,'') + ISNULL(P.MEMO,'') AS KEY_NAME, 
                    P.REMARK, P.MEMO, P.MESSAGE,
                    CHILD.PR_ID AS CHILD_PR_ID, 
                    P.PR_ID AS Dminus3day, 
                    P.PARENT_PR_ID AS PARENT_PR_ID
                    FROM PLAN_PURCHASE_REQUISITION P
                    LEFT JOIN PLAN_PURCHASE_REQUISITION CHILD
                    ON P.PR_ID=CHILD.PARENT_PR_ID
                    WHERE P.PROJECT_ID=@projectid 
                    AND P.PR_ID Like Concat(@type,'%') ";
            StringBuilder sb = new StringBuilder(sql);
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("type", type));
            parameters.Add(new SqlParameter("projectid", projectid));
            //加入狀態條件
            if (null != status && "*" != status)
            {
                sb.Append(" AND P.STATUS =@status");
                parameters.Add(new SqlParameter("status", status));
            }
            //加入申請年月條件
            if (null != yymm && "" != yymm)
            {
                string[] period = yymm.Split('/');
                string Year = period[0];
                string Month = period[1];
                sb.Append(" AND YEAR(P.CREATE_DATE) = @year AND MONTH(P.CREATE_DATE) = @month");
                parameters.Add(new SqlParameter("year", Year));
                parameters.Add(new SqlParameter("month", Month));
            }
            //加入單號條件
            if (null != prid && "" != prid)
            {
                sb.Append(" AND P.PR_ID=@prid ");
                parameters.Add(new SqlParameter("prid", prid));
            }
            //加入關鍵字
            if (null != keyword && "" != keyword)
            {

                sb.Append(" AND (P.REMARK LIKE '%' + @keyword + '%' OR P.MEMO '%' + @keyword + '%'  ");
                parameters.Add(new SqlParameter("keyword", keyword));
            }
            sql = sb.ToString();
            using (var context = new topmepEntities())
            {
                logger.Debug("get purchase requisition sql=" + sql);
                lstForm = context.Database.SqlQuery<PRFunction>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get purchase requisition count=" + lstForm.Count);
            return lstForm;
        }
        //取得申購單資料
        public List<PRFunction> getPRByPrjId(string projectid, string date, string keyname, string prid, string status)
        {
            logger.Info("search purchase requisition by 申購日期 =" + date + ", 申購單編號 =" + prid + ", 關鍵字 =" + keyname + ", 申購單狀態 =" + status);
            //(string type, string projectid, string status, string yymm, string prid,string keyword)
            List<PRFunction> lstForm = getPRFunction("PR", projectid, status, date, prid, keyname);
            return lstForm;
        }

        //取得申購單
        public void getPRByPrId(string prid, string parentId, string prjid)
        {
            logger.Info("get form : formid=" + prid);
            using (var context = new topmepEntities())
            {
                //取得申購單檔頭資訊
                string sql = "SELECT PR_ID, PROJECT_ID, RECIPIENT, LOCATION, PRJ_UID, CREATE_USER_ID, CREATE_DATE, REMARK, SUPPLIER_ID, MODIFY_DATE, PARENT_PR_ID, STATUS, MEMO, MESSAGE, CAUTION FROM " +
                    "PLAN_PURCHASE_REQUISITION WHERE PR_ID =@prid ";
                formPR = context.PLAN_PURCHASE_REQUISITION.SqlQuery(sql, new SqlParameter("prid", prid)).First();
                //取得申購單明細
                string sql4Po = @"SELECT pri.NEED_QTY, CONVERT(char(10), pri.NEED_DATE, 111) AS NEED_DATE, 
                                pri.REMARK, pri.PR_ITEM_ID, pri.ORDER_QTY, pri.PLAN_ITEM_ID, pri.RECEIPT_QTY, pi.ITEM_ID, 
                                IIF(pi.ITEM_DESC IS NOT NULL, pi.ITEM_DESC, pri.ITEM_DESC) AS ITEM_DESC, 
                                IIF(pi.ITEM_UNIT IS NOT NULL, pi.ITEM_UNIT, pri.ITEM_UNIT) AS ITEM_UNIT, 
                                pi.ITEM_FORM_QUANTITY, pi.SYSTEM_MAIN, md.QTY AS MAP_QTY,  
                                B.CUMULATIVE_QTY, C.RECEIPT_QTY_BY_PO, E.ALL_RECEIPT_QTY, 
                                E.ALL_RECEIPT_QTY - D.DELIVERY_QTY AS INVENTORY_QTY,
                                ROW_NUMBER() OVER(ORDER BY pi.EXCEL_ROW_ID) AS NO 
                                FROM PLAN_PURCHASE_REQUISITION_ITEM pri 
                                LEFT JOIN PLAN_ITEM pi ON pri.PLAN_ITEM_ID = pi.PLAN_ITEM_ID 
                                LEFT JOIN vw_MAP_MATERLIALIST md ON pi.PLAN_ITEM_ID = md.PROJECT_ITEM_ID 
                                LEFT JOIN (
                                SELECT pri.PLAN_ITEM_ID, SUM(pri.ORDER_QTY) AS CUMULATIVE_QTY FROM PLAN_PURCHASE_REQUISITION_ITEM pri 
                                LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pri.PR_ID = pr.PR_ID 
                                WHERE pr.PROJECT_ID =@prjid AND pri.PR_ID LIKE 'PPO%' GROUP BY pri.PLAN_ITEM_ID
                                ) B ON pri.PLAN_ITEM_ID = B.PLAN_ITEM_ID 
                                LEFT JOIN(
                                SELECT pri.PLAN_ITEM_ID, SUM(pri.RECEIPT_QTY) AS RECEIPT_QTY_BY_PO 
                                FROM PLAN_PURCHASE_REQUISITION_ITEM pri 
                                LEFT JOIN PLAN_PURCHASE_REQUISITION ppr ON pri.PR_ID = ppr.PR_ID 
                                WHERE ppr.PARENT_PR_ID =@parentId GROUP BY pri.PLAN_ITEM_ID
                                )C ON pri.PLAN_ITEM_ID = C.PLAN_ITEM_ID 
                                LEFT JOIN (
                                SELECT pri.PLAN_ITEM_ID, SUM(pri.RECEIPT_QTY) AS ALL_RECEIPT_QTY 
                                FROM PLAN_PURCHASE_REQUISITION_ITEM pri 
                                LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pri.PR_ID = pr.PR_ID 
                                WHERE pr.PROJECT_ID = @prjid AND pri.PR_ID LIKE 'RP%' GROUP BY pri.PLAN_ITEM_ID
                                )E ON pri.PLAN_ITEM_ID = E.PLAN_ITEM_ID LEFT JOIN (
                                SELECT pid.PLAN_ITEM_ID, SUM(pid.DELIVERY_QTY) AS DELIVERY_QTY FROM PLAN_ITEM_DELIVERY pid 
                                LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pid.DELIVERY_ORDER_ID = pr.PR_ID 
                                WHERE pr.PROJECT_ID = @prjid GROUP BY pid.PLAN_ITEM_ID
                                )D ON pri.PLAN_ITEM_ID = D.PLAN_ITEM_ID WHERE pri.PR_ID =@prid
";
                logger.Debug("get purchase requisition item :" + sql4Po);
                PRItem = context.Database.SqlQuery<PurchaseRequisition>(sql4Po, new SqlParameter("prid", prid), new SqlParameter("parentId", parentId), new SqlParameter("prjid", prjid)).ToList();

                //取得領料明細
                string sql4Deliver = @"SELECT pid.PLAN_ITEM_ID, pid.REMARK, pid.DELIVERY_QTY, 
                    pi.ITEM_ID, pi.ITEM_DESC, pi.ITEM_UNIT, pi.SYSTEM_MAIN, 
                    ROW_NUMBER() OVER(ORDER BY pi.EXCEL_ROW_ID) AS NO 
                    FROM PLAN_ITEM_DELIVERY pid 
                    LEFT JOIN PLAN_ITEM pi ON pid.PLAN_ITEM_ID = pi.PLAN_ITEM_ID 
                    WHERE pid.DELIVERY_ORDER_ID =@prid";
                logger.Debug("get purchase requisition item :" + sql4Po + ",prid=" + prid);
                DOItem = context.Database.SqlQuery<PurchaseRequisition>(sql4Deliver, new SqlParameter("prid", prid)).ToList();

                logger.Debug("get delivery item count:" + DOItem.Count);
            }
        }

        //新增申購單物料品項
        public int addPRItem(PLAN_PURCHASE_REQUISITION_ITEM item)
        {
            int i = 0;
            //2.將資料寫入 
            using (var context = new topmepEntities())
            {
                try
                {
                    logger.Debug("INSERT ITEM TO PLAN_PURCHASE_REQUISITION:" + item.PR_ID);
                    context.PLAN_PURCHASE_REQUISITION_ITEM.Add(item);
                    i = context.SaveChanges();
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":" + ex.StackTrace);
                }
            }
            return i;
        }
        public string getParentPrIdByPrId(string prid)
        {
            string parentid = null;
            using (var context = new topmepEntities())
            {
                parentid = context.Database.SqlQuery<string>("select DISTINCT PARENT_PR_ID FROM PLAN_PURCHASE_REQUISITION WHERE PARENT_PR_ID =@prid  "
               , new SqlParameter("prid", prid)).FirstOrDefault();
            }
            return parentid;
        }
        public PLAN_PURCHASE_REQUISITION table = null;
        //更新申購單資料
        public int updatePR(string formid, PLAN_PURCHASE_REQUISITION pr, List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem, SYS_USER u)
        {
            logger.Info("Update purchase requisition id =" + formid);
            table = pr;
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.Entry(table).State = EntityState.Modified;
                    i = context.SaveChanges();
                    logger.Debug("Update purchase requisition =" + i);
                    logger.Info("purchase requisition item = " + lstItem.Count);
                    //2.將item資料寫入 
                    foreach (PLAN_PURCHASE_REQUISITION_ITEM item in lstItem)
                    {
                        PLAN_PURCHASE_REQUISITION_ITEM existItem = null;
                        logger.Debug("purchase requisition item id=" + item.PR_ITEM_ID);
                        if (item.PR_ITEM_ID != 0)
                        {
                            existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(item.PR_ITEM_ID);
                        }
                        else
                        {
                            var parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("formid", formid));
                            parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                            string sql = "SELECT * FROM PLAN_PURCHASE_REQUISITION_ITEM WHERE PR_ID=@formid AND PLAN_ITEM_ID=@itemid";
                            logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                            PLAN_PURCHASE_REQUISITION_ITEM excelItem = context.PLAN_PURCHASE_REQUISITION_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                            existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(excelItem.PR_ITEM_ID);

                        }
                        logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                        existItem.NEED_QTY = item.NEED_QTY;
                        existItem.NEED_DATE = item.NEED_DATE;
                        existItem.REMARK = item.REMARK;
                        context.PLAN_PURCHASE_REQUISITION_ITEM.AddOrUpdate(existItem);
                    }
                    ///
                    i = i + context.SaveChanges();
                    ///send email to  業管
                    if (u != null)
                    {
                        EMailService email = new EMailService();
                        email.createPRMessage(u, pr);
                    }
                    logger.Debug("Update purchase requisition item =" + i);
                }
                catch (Exception e)
                {
                    logger.Error("update new purchase requisition id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }
            }
            return i;
        }
        /// <summary>
        /// 採購單作業資料:將申購單依據發包廠商建立不同記錄
        /// </summary>
        /// <param name="projectid"></param>
        /// <returns></returns>
        public List<PurchaseOrderFunction> getPRBySupplier(string projectid)
        {
            List<PurchaseOrderFunction> lstPO = new List<PurchaseOrderFunction>();
            using (var context = new topmepEntities())
            {
                string sql = @"SELECT CONVERT(char(10), B.CREATE_DATE, 111) AS CREATE_DATE, B.PR_ID, 
                ISNULL(B.SUPPLIER_ID, '') AS SUPPLIER_ID, 
                MIN(CONVERT(char(10), B.NEED_DATE, 111)) AS NEED_DATE, 
                B.PROJECT_ID 
                FROM (
                SELECT DISTINCT(A.PROJECT_ID + '-' + PR_ID + '-' + ISNULL(SUPPLIER_ID,'') + '-' + CONVERT(char(10), NEED_DATE, 111)) AS NAME, 
                CONVERT(char(10), A.CREATE_DATE, 111) AS CREATE_DATE, A.PR_ID, A.SUPPLIER_ID, A.PROJECT_ID, A.NEED_DATE 
                FROM (
                SELECT pri.*, pi.ITEM_ID, pi.SUPPLIER_ID, pr.CREATE_DATE, pr.PROJECT_ID 
                FROM PLAN_PURCHASE_REQUISITION_ITEM pri 
                JOIN PLAN_ITEM pi ON pri.PLAN_ITEM_ID = pi.PLAN_ITEM_ID 
                LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pri.PR_ID = pr.PR_ID 
                AND ISNULL(pi.SUPPLIER_ID,'') != ''
                WHERE pr.PROJECT_ID =@projectid AND pr.SUPPLIER_ID IS NULL AND pr.STATUS > 5 AND pr.PR_ID LIKE 'PR%'
                ) A 
                WHERE A.PR_ID + ISNULL(A.SUPPLIER_ID,'') NOT IN (SELECT DISTINCT(pr.PARENT_PR_ID + pr.SUPPLIER_ID) AS ORDER_RECORD 
                FROM PLAN_PURCHASE_REQUISITION pr WHERE pr.PR_ID LIKE 'PPO%'))B 
                GROUP BY CONVERT(char(10), B.CREATE_DATE, 111), B.PR_ID, B.SUPPLIER_ID, B.PROJECT_ID ORDER BY NEED_DATE ";

                logger.Info("sql = " + sql);
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", projectid));
                lstPO = context.Database.SqlQuery<PurchaseOrderFunction>(sql, parameters.ToArray()).ToList();
                logger.Info("Get Purchase Requisition By Suplier Record Count =" + lstPO.Count);
            }
            return lstPO;
        }
        //取得申購單項目by供應商
        public List<PurchaseRequisition> getPurchaseItemBySupplier(string id, string prjid)
        {
            //取得各供應商採購內容
            logger.Info("get purchase requisition item by supplier ");
            List<PurchaseRequisition> lstItem = new List<PurchaseRequisition>();
            using (var context = new topmepEntities())
            {
                string sql = "SELECT pi.PLAN_ITEM_ID, pri.PR_ITEM_ID, pri.NEED_QTY, CONVERT(char(10), pri.NEED_DATE, 111) AS NEED_DATE, pri.REMARK , pi.ITEM_ID, pi.ITEM_DESC, pi.ITEM_UNIT, pi.ITEM_FORM_QUANTITY, " +
                    "pi.SUPPLIER_ID, B.CUMULATIVE_QTY FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_ITEM pi on pri.PLAN_ITEM_ID = pi.PLAN_ITEM_ID " +
                    "LEFT JOIN (SELECT pri.PLAN_ITEM_ID, pri.REMARK, SUM(pri.ORDER_QTY) AS CUMULATIVE_QTY FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_PURCHASE_REQUISITION pr " +
                    "ON pri.PR_ID = pr.PR_ID WHERE pr.PROJECT_ID =@prjid AND pri.PR_ID LIKE 'PR%' GROUP BY pri.PLAN_ITEM_ID, pri.REMARK)B " +
                    "ON pri.PLAN_ITEM_ID + pri.REMARK = B.PLAN_ITEM_ID + B.REMARK WHERE pri.PR_ID + '-' + ISNULL(pi.SUPPLIER_ID,'') =@id ";

                logger.Info("sql = " + sql);
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("id", id));
                parameters.Add(new SqlParameter("prjid", prjid));
                lstItem = context.Database.SqlQuery<PurchaseRequisition>(sql, parameters.ToArray()).ToList();
                logger.Info("Get purchase requisition item by supplier Record Count=" + lstItem.Count);
            }
            return lstItem;
        }
        //更新申購單狀態為退件不處理(狀態代碼為5)
        public int changePRStatus(string formid, string message)
        {
            int i = 0;
            logger.Info("Update PR Status and Message, it's formid=" + formid + ", messaage =" + message);
            db = new topmepEntities();
            string sql = "UPDATE PLAN_PURCHASE_REQUISITION SET STATUS = 0, MESSAGE=@message, MODIFY_DATE =@datetime  WHERE PR_ID=@formid ";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("message", message));
            parameters.Add(new SqlParameter("formid", formid));
            parameters.Add(new SqlParameter("datetime", DateTime.Now));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            db = null;
            logger.Debug("Update PR Status :" + i);
            return i;
        }
        //更新申購單狀態為退件(狀態代碼為-10)
        public int RejectPRByPRId(string prid)
        {
            int i = 0;
            logger.Info("reject PR form by prid" + prid);
            using (var context = new topmepEntities())
            {
                UpdateStatus(prid, null, "-10", context);
            }
            return 1;
        }
        //更新申購單/採購單/驗收單狀態
        protected void UpdateStatus(string prid, string parent_id, string status, topmepEntities context)
        {
            string sql = null;
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("status", status));
            if (null != prid && "" != prid)
            {
                sql = "UPDATE  PLAN_PURCHASE_REQUISITION SET STATUS = @status WHERE PR_ID = @prid ";
                parameters.Add(new SqlParameter("prid", prid));
            }
            else
            {
                sql = "UPDATE  PLAN_PURCHASE_REQUISITION SET STATUS = @status WHERE PARENT_PR_ID = @parent_id ";
                parameters.Add(new SqlParameter("parent_id", parent_id));
            }
            logger.Debug("sql=" + sql + ",status=" + status + ",prid=" + prid + ",parend_pr_id" + parent_id);
            context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            int i = context.SaveChanges();
            logger.Debug("Update Count=" + i);
        }
        // 寫入採購內容
        public string newPO(string projectid, PLAN_PURCHASE_REQUISITION form, string[] lstItemId, string parentid)
        {
            //1.建立採購單
            logger.Info("create new purchase order ");
            string sno_key = "PPO";
            SerialKeyService snoservice = new SerialKeyService();
            form.PR_ID = snoservice.getSerialKey(sno_key);
            logger.Info("new purchase order =" + form.ToString());
            using (var context = new topmepEntities())
            {
                context.PLAN_PURCHASE_REQUISITION.Add(form);
                int i = context.SaveChanges();
                logger.Debug("Add Purchase Order=" + i);
                logger.Info("plan purchase Order id = " + form.PR_ID);
                //if (i > 0) { status = true; };
                List<topmeperp.Models.PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
                string ItemId = "";
                for (i = 0; i < lstItemId.Count(); i++)
                {
                    if (i < lstItemId.Count() - 1)
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'" + ",";
                    }
                    else
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'";
                    }
                }

                string sql = "INSERT INTO PLAN_PURCHASE_REQUISITION_ITEM (PR_ID, PLAN_ITEM_ID, NEED_QTY, NEED_DATE, REMARK) "
                + "SELECT '" + form.PR_ID + "' as PR_ID, A.PLAN_ITEM_ID as PLAN_ITEM_ID, A.NEED_QTY as NEED_QTY, A.NEED_DATE as NEED_DATE, A.REMARK as REMARK  "
                + "FROM (SELECT pri.* FROM PLAN_PURCHASE_REQUISITION_ITEM pri WHERE pri.PR_ID = '" + parentid + "' AND pri.PLAN_ITEM_ID IN (" + ItemId + "))A ";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                //更新申購單狀態
                ////UpdateStatus(parentid, null, "40", context);
                return form.PR_ID;
            }
        }
        //更新採購數量
        public int refreshPO(string formid, PLAN_PURCHASE_REQUISITION form, List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem, SYS_USER u)
        {
            logger.Info("Update plan purchase order id =" + formid);
            //int i = 0;
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.Entry(form).State = EntityState.Modified;
                    i = context.SaveChanges();
                    logger.Debug("Update plan purchase order =" + i);
                    logger.Info("purchase order item = " + lstItem.Count);
                    //2.將item資料寫入 
                    foreach (PLAN_PURCHASE_REQUISITION_ITEM item in lstItem)
                    {
                        PLAN_PURCHASE_REQUISITION_ITEM existItem = null;
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("formid", formid));
                        parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                        string sql = "SELECT * FROM PLAN_PURCHASE_REQUISITION_ITEM WHERE PR_ID=@formid AND PLAN_ITEM_ID=@itemid";
                        logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                        PLAN_PURCHASE_REQUISITION_ITEM excelItem = context.PLAN_PURCHASE_REQUISITION_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(excelItem.PR_ITEM_ID);
                        logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                        existItem.ORDER_QTY = item.ORDER_QTY;
                        context.PLAN_PURCHASE_REQUISITION_ITEM.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    ///send email to  工地主任與申請人
                    EMailService email = new EMailService();
                    email.createPOMessage(u, form);

                    logger.Debug("Update purchase order item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new purchase order id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }

        //取得採購單資料
        public List<PRFunction> getPOByPrjId(string projectid, string date, string supplier, string prid, string parentPrid, string keyname)
        {

            logger.Info("search purchase order by 採購日期 =" + date + ", 採購單編號 =" + prid + ", 供應商名稱 =" + supplier + ", 申購單編號 =" + parentPrid + ", 關鍵字 =" + keyname);
            List<PRFunction> lstForm = new List<PRFunction>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT CONVERT(char(10), CREATE_DATE, 111) AS CREATE_DATE, PR_ID, PROJECT_ID, SUPPLIER_ID, PARENT_PR_ID, " +
                "ISNULL(CAUTION,'') + ISNULL(MEMO,'') + ISNULL(MESSAGE,'') AS KEY_NAME, ROW_NUMBER() OVER(ORDER BY PR_ID) AS NO " +
                "FROM PLAN_PURCHASE_REQUISITION WHERE PROJECT_ID =@projectid AND SUPPLIER_ID IS NOT NULL AND PR_ID NOT LIKE 'RP%' ORDER BY STATUS ";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectid));
            //採購年月查詢條件
            if (null != date && date != "")
            {
                //DateTime dt = Convert.ToDateTime(date);
                //string DateString = dt.AddDays(1).ToString("yyyy/MM/dd");
                //sql = sql + "AND CREATE_DATE >=@date AND  CREATE_DATE < '" + DateString + "' ";
                string[] period = date.Split('/');
                string Year = period[0];
                string Month = period[1];
                sql = sql + "AND YEAR(CREATE_DATE) = '" + Year + "' AND MONTH(CREATE_DATE) = '" + Month + "' ";
                parameters.Add(new SqlParameter("date", date));
            }
            //關鍵字條件
            if (null != keyname && keyname != "")
            {
                sql = sql + "AND ISNULL(CAUTION,'') + ISNULL(MEMO,'') + ISNULL(MESSAGE,'') LIKE @keyname ";
                parameters.Add(new SqlParameter("keyname", '%' + keyname + '%'));
            }
            //採購單編號條件
            if (null != prid && prid != "")
            {
                sql = sql + "AND PR_ID =@prid ";
                parameters.Add(new SqlParameter("prid", prid));
            }
            //供應商條件
            if (null != supplier && supplier != "")
            {
                sql = sql + "AND SUPPLIER_ID LIKE @supplier ";
                parameters.Add(new SqlParameter("supplier", '%' + supplier + '%'));
            }
            //申購單編號條件
            if (null != parentPrid && parentPrid != "")
            {
                sql = sql + "AND PARENT_PR_ID =@parentPrid ";
                parameters.Add(new SqlParameter("parentPrid", parentPrid));
            }
            using (var context = new topmepEntities())
            {
                logger.Debug("get purchase order sql=" + sql);
                lstForm = context.Database.SqlQuery<PRFunction>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get purchase order count=" + lstForm.Count);
            return lstForm;
        }

        //更新採購單資料
        public int updatePO(string formid, PLAN_PURCHASE_REQUISITION pr, List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem)
        {
            logger.Info("Update purchase order id =" + formid);
            table = pr;
            int i = 0;
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.Entry(table).State = EntityState.Modified;
                    i = context.SaveChanges();
                    logger.Debug("Update purchase order =" + i);
                    logger.Info("purchase order item = " + lstItem.Count);
                    //2.將item資料寫入 
                    foreach (PLAN_PURCHASE_REQUISITION_ITEM item in lstItem)
                    {
                        PLAN_PURCHASE_REQUISITION_ITEM existItem = null;
                        logger.Debug("purchase order item id=" + item.PR_ITEM_ID);
                        if (item.PR_ITEM_ID != 0)
                        {
                            existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(item.PR_ITEM_ID);
                        }
                        else
                        {
                            var parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("formid", formid));
                            parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                            string sql = "SELECT * FROM PLAN_PURCHASE_REQUISITION_ITEM WHERE PR_ID=@formid AND PLAN_ITEM_ID=@itemid";
                            logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                            PLAN_PURCHASE_REQUISITION_ITEM excelItem = context.PLAN_PURCHASE_REQUISITION_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                            existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(excelItem.PR_ITEM_ID);

                        }
                        logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                        existItem.ORDER_QTY = item.ORDER_QTY;
                        context.PLAN_PURCHASE_REQUISITION_ITEM.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update purchase order item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new purchase order id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }
        //更新採購單memo
        public int changeMemo(string formid, string memo)
        {
            int i = 0;
            logger.Info("Update PO memo, it's formid=" + formid + ", memo =" + memo);
            db = new topmepEntities();
            string sql = "UPDATE PLAN_PURCHASE_REQUISITION SET MEMO=@memo, MODIFY_DATE =@datetime  WHERE PR_ID=@formid ";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("memo", memo));
            parameters.Add(new SqlParameter("formid", formid));
            parameters.Add(new SqlParameter("datetime", DateTime.Now));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            db = null;
            logger.Debug("Update PO memo :" + i);
            return i;
        }
        // 寫入驗收內容
        public string newRP(string projectid, PLAN_PURCHASE_REQUISITION form, string[] lstItemId, string parentid)
        {
            //1.建立驗收資料
            logger.Info("create new receipt ");
            string sno_key = "RP";
            SerialKeyService snoservice = new SerialKeyService();
            form.PR_ID = snoservice.getSerialKey(sno_key);
            logger.Info("new purchase receipt =" + form.ToString());
            using (var context = new topmepEntities())
            {
                context.PLAN_PURCHASE_REQUISITION.Add(form);
                int i = context.SaveChanges();
                logger.Debug("Add Purchase receipt=" + i);
                logger.Info("plan purchase receipt id = " + form.PR_ID);
                //if (i > 0) { status = true; };
                List<topmeperp.Models.PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
                string ItemId = "";
                for (i = 0; i < lstItemId.Count(); i++)
                {
                    if (i < lstItemId.Count() - 1)
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'" + ",";
                    }
                    else
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'";
                    }
                }

                string sql = "INSERT INTO PLAN_PURCHASE_REQUISITION_ITEM (PR_ID, PLAN_ITEM_ID, NEED_QTY, NEED_DATE, REMARK, ORDER_QTY) "
                + "SELECT '" + form.PR_ID + "' as PR_ID, A.PLAN_ITEM_ID as PLAN_ITEM_ID, A.NEED_QTY as NEED_QTY, A.NEED_DATE as NEED_DATE, A.REMARK as REMARK, A.ORDER_QTY as ORDER_QTY  "
                + "FROM (SELECT pri.* FROM PLAN_PURCHASE_REQUISITION_ITEM pri WHERE pri.PR_ID = '" + parentid + "' AND pri.PLAN_ITEM_ID IN (" + ItemId + "))A ";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return form.PR_ID;
            }
        }

        //更新驗收數量
        public int refreshRP(string formid, PLAN_PURCHASE_REQUISITION form, List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem, string closeFlag)
        {
            logger.Info("Update plan purchase receipt id =" + formid);
            int i = 0;
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.Entry(form).State = EntityState.Modified;
                    i = context.SaveChanges();
                    logger.Debug("Update plan purchase receipt =" + i);
                    logger.Info("purchase receipt item = " + lstItem.Count);
                    //2.將item資料寫入 
                    foreach (PLAN_PURCHASE_REQUISITION_ITEM item in lstItem)
                    {
                        PLAN_PURCHASE_REQUISITION_ITEM existItem = null;
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("formid", formid));
                        parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                        string sql = "SELECT * FROM PLAN_PURCHASE_REQUISITION_ITEM WHERE PR_ID=@formid AND PLAN_ITEM_ID=@itemid";
                        logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                        PLAN_PURCHASE_REQUISITION_ITEM excelItem = context.PLAN_PURCHASE_REQUISITION_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(excelItem.PR_ITEM_ID);
                        logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                        existItem.RECEIPT_QTY = item.RECEIPT_QTY;
                        context.PLAN_PURCHASE_REQUISITION_ITEM.AddOrUpdate(existItem);
                        //更新採購單狀態
                        if (closeFlag == "Y")
                        {
                            UpdateStatus(form.PARENT_PR_ID, null, "40", context);
                        }
                        else
                        {
                            UpdateStatus(form.PARENT_PR_ID, null, "30", context);
                        }
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update purchase reeipt item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new purchase receipt id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }

        //更新驗收單資料//採購單結案
        //public int updateRP(string formid, PLAN_PURCHASE_REQUISITION pr, List<PLAN_PURCHASE_REQUISITION_ITEM> lstItem)
        //{
        //    logger.Info("Update purchase receipt id =" + formid);
        //    table = pr;
        //    int i = 0;
        //    int j = 0;
        //    using (var context = new topmepEntities())
        //    {
        //        try
        //        {
        //            context.Entry(table).State = EntityState.Modified;
        //            i = context.SaveChanges();
        //            logger.Debug("Update purchase receipt =" + i);
        //            logger.Info("purchase receipt item = " + lstItem.Count);
        //            //2.將item資料寫入 
        //            foreach (PLAN_PURCHASE_REQUISITION_ITEM item in lstItem)
        //            {
        //                PLAN_PURCHASE_REQUISITION_ITEM existItem = null;
        //                logger.Debug("purchase receipt item id=" + item.PR_ITEM_ID);
        //                if (item.PR_ITEM_ID != 0)
        //                {
        //                    existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(item.PR_ITEM_ID);
        //                }
        //                else
        //                {
        //                    var parameters = new List<SqlParameter>();
        //                    parameters.Add(new SqlParameter("formid", formid));
        //                    parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
        //                    string sql = "SELECT * FROM PLAN_PURCHASE_REQUISITION_ITEM WHERE PR_ID=@formid AND PLAN_ITEM_ID=@itemid";
        //                    logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
        //                    PLAN_PURCHASE_REQUISITION_ITEM excelItem = context.PLAN_PURCHASE_REQUISITION_ITEM.SqlQuery(sql, parameters.ToArray()).First();
        //                    existItem = context.PLAN_PURCHASE_REQUISITION_ITEM.Find(excelItem.PR_ITEM_ID);

        //                }
        //                logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
        //                existItem.RECEIPT_QTY = item.RECEIPT_QTY;
        //                context.PLAN_PURCHASE_REQUISITION_ITEM.AddOrUpdate(existItem);
        //            }
        //            //更新採購單狀態
        //            UpdateStatus(pr.PARENT_PR_ID, null, "30", context);
        //            j = context.SaveChanges();
        //            logger.Debug("Update purchase receipt item =" + j);
        //            return j;
        //        }
        //        catch (Exception e)
        //        {
        //            logger.Error("update new purchase receipt id fail:" + e.ToString());
        //            logger.Error(e.StackTrace);
        //            message = e.Message;
        //        }

        //    }
        //    return i;
        //}

        //取得驗收單資料
        public List<PRFunction> getRPByPrId(string prid)
        {

            logger.Info("search purchase receipt by 採購單編號 =" + prid);
            List<PRFunction> lstForm = new List<PRFunction>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT CONVERT(char(10), CREATE_DATE, 111) AS CREATE_DATE, PR_ID, SUPPLIER_ID, PARENT_PR_ID, ROW_NUMBER() OVER(ORDER BY PR_ID) AS NO " +
                "FROM PLAN_PURCHASE_REQUISITION WHERE SUPPLIER_ID IS NOT NULL AND PARENT_PR_ID =@prid AND PR_ID LIKE 'RP%' ";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("prid", prid));

            using (var context = new topmepEntities())
            {
                logger.Debug("get purchase receipt sql=" + sql);
                lstForm = context.Database.SqlQuery<PRFunction>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get purchase receipt count=" + lstForm.Count);
            return lstForm;
        }

        //取得驗收單資料以供領料使用
        public List<PRFunction> getRP4Delivery(string prjid, string keyword)
        {

            logger.Info("search receipt for delivery by 專案編號 =" + prjid + ", 關鍵字名稱 =" + keyword);
            List<PRFunction> lstForm = new List<PRFunction>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT CONVERT(char(10), CREATE_DATE, 111) AS CREATE_DATE, PR_ID, SUPPLIER_ID, REMARK, MEMO, MESSAGE, ROW_NUMBER() OVER(ORDER BY PR_ID) AS NO " +
                "FROM PLAN_PURCHASE_REQUISITION WHERE PROJECT_ID =@prjid AND PR_ID LIKE 'RP%' ";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("prjid", prjid));

            //關鍵字條件
            if (null != keyword && keyword != "")
            {
                sql = sql + "AND MEMO LIKE @keyword OR PROJECT_ID =@prjid AND PR_ID LIKE 'RP%' AND MESSAGE LIKE @keyword OR PROJECT_ID =@prjid AND PR_ID LIKE 'RP%' AND REMARK LIKE @keyword ";
                parameters.Add(new SqlParameter("keyword", '%' + keyword + '%'));
            }
            using (var context = new topmepEntities())
            {
                logger.Debug("get receipt for delivery sql=" + sql);
                lstForm = context.Database.SqlQuery<PRFunction>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get receipt for delivery count=" + lstForm.Count);
            return lstForm;
        }
        //取得物料庫存數量
        public List<PurchaseRequisition> getInventoryByPrjId(string prjid, string itemName, string typeMain, string typeSub, string systemMain, string systemSub, string remark)
        {

            logger.Info("search inventory by 專案編號 =" + prjid + ", 物料名稱 =" + itemName + ", 九宮格 =" + typeMain + ", 次九宮格 =" + typeSub + ", 主系統 =" + systemMain + ", 次系統 =" + systemSub + ", 備註 =" + remark);
            List<PurchaseRequisition> lstItem = new List<PurchaseRequisition>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT pri.PLAN_ITEM_ID, pri.REMARK, pr.PROJECT_ID, pi.ITEM_ID, pi.ITEM_DESC, pi.ITEM_UNIT, pi.SYSTEM_MAIN, SUM(pri.RECEIPT_QTY) - ISNULL(A.DELIVERY_QTY, 0) AS INVENTORY_QTY, B.diffQty " +
                "FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pri.PR_ID = pr.PR_ID LEFT JOIN PLAN_ITEM pi ON pri.PLAN_ITEM_ID = pi.PLAN_ITEM_ID " +
                "LEFT JOIN (SELECT pid.PLAN_ITEM_ID, pid.REMARK, SUM(pid.DELIVERY_QTY) AS DELIVERY_QTY FROM PLAN_ITEM_DELIVERY pid " +
                "GROUP BY pid.PLAN_ITEM_ID, pid.REMARK)A ON pri.PLAN_ITEM_ID + pri.REMARK = A.PLAN_ITEM_ID + A.REMARK LEFT JOIN " +
                "(SELECT pri.PLAN_ITEM_ID, pri.REMARK, SUM(ISNULL(pri.NEED_QTY, 0)) - ISNULL(main.RECEIPT_QTY, 0) AS diffQty " +
                "FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pri.PR_ID = pr.PR_ID LEFT JOIN " +
                "(SELECT pri.PLAN_ITEM_ID, pri.REMARK, SUM(pri.RECEIPT_QTY)AS RECEIPT_QTY FROM PLAN_PURCHASE_REQUISITION_ITEM pri " +
                "LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pri.PR_ID = pr.PR_ID WHERE pr.PROJECT_ID =@prjid AND pri.PR_ID LIKE 'RP%' " +
                "GROUP BY pri.PLAN_ITEM_ID, pri.REMARK)main ON main.PLAN_ITEM_ID + main.REMARK = pri.PLAN_ITEM_ID + pri.REMARK " +
                "WHERE pr.PROJECT_ID =@prjid AND pri.PR_ID LIKE 'PPO%' GROUP BY pri.PLAN_ITEM_ID, pri.REMARK, main.RECEIPT_QTY)B ON pri.PLAN_ITEM_ID + pri.REMARK = B.PLAN_ITEM_ID + B.REMARK " +
                "WHERE pr.PR_ID NOT LIKE 'PR%' GROUP BY pri.PLAN_ITEM_ID, pri.REMARK, pr.PROJECT_ID, A.DELIVERY_QTY, B.diffQty, " +
                "pi.ITEM_ID, pi.ITEM_DESC, pi.ITEM_UNIT, pi.TYPE_CODE_1, pi.TYPE_CODE_2, pi.SYSTEM_MAIN, pi.SYSTEM_SUB, pi.EXCEL_ROW_ID HAVING pr.PROJECT_ID =@prjid ";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("prjid", prjid));

            //物料名稱條件
            if (null != itemName && itemName != "")
            {
                sql = sql + "AND pi.ITEM_DESC LIKE @itemName ";
                parameters.Add(new SqlParameter("itemName", '%' + itemName + '%'));
            }
            //項目備註條件
            if (null != remark && remark != "")
            {
                sql = sql + "AND pri.REMARK LIKE @remark ";
                parameters.Add(new SqlParameter("remark", '%' + remark + '%'));
            }
            //九宮格條件
            if (null != typeMain && typeMain != "")
            {
                sql = sql + "AND pi.TYPE_CODE_1 =@typeMain ";
                parameters.Add(new SqlParameter("typeMain", typeMain));
            }
            //次九宮格條件
            if (null != typeSub && typeSub != "")
            {
                sql = sql + "AND pi.TYPE_CODE_2 =@typeSub ";
                parameters.Add(new SqlParameter("typeSub", typeSub));
            }
            //主系統條件
            if (null != systemMain && systemMain != "")
            {
                sql = sql + "AND REPLACE(pi.SYSTEM_MAIN,' ','') =@systemMain ";
                parameters.Add(new SqlParameter("systemMain", systemMain));
            }
            //次系統條件
            if (null != systemSub && systemSub != "")
            {
                sql = sql + "AND REPLACE(pi.SYSTEM_SUB,' ','') =@systemSub ";
                parameters.Add(new SqlParameter("systemSub", systemSub));
            }
            sql = sql + "ORDER BY pi.EXCEL_ROW_ID ASC ";
            using (var context = new topmepEntities())
            {
                logger.Debug("get inventory sql=" + sql);
                lstItem = context.Database.SqlQuery<PurchaseRequisition>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get inventory count=" + lstItem.Count);
            return lstItem;
        }

        // 寫入領料內容
        public string newDelivery(string projectid, PLAN_PURCHASE_REQUISITION form, string[] lstItemId, string createid)
        {
            //1.新增領料品項
            logger.Info("create new delivery item ");
            string sno_key = "DO";
            SerialKeyService snoservice = new SerialKeyService();
            form.PR_ID = snoservice.getSerialKey(sno_key);
            logger.Info("new delivery form =" + form.ToString());
            using (var context = new topmepEntities())
            {
                context.PLAN_PURCHASE_REQUISITION.Add(form);
                int i = context.SaveChanges();
                logger.Debug("Add Delivery Form=" + i);
                logger.Info("plan delivery form id = " + form.PR_ID);
                //if (i > 0) { status = true; };
                List<topmeperp.Models.PLAN_ITEM_DELIVERY> lstItem = new List<PLAN_ITEM_DELIVERY>();
                string ItemId = "";
                for (i = 0; i < lstItemId.Count(); i++)
                {
                    if (i < lstItemId.Count() - 1)
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'" + ",";
                    }
                    else
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'";
                    }
                }

                string sql = "INSERT INTO PLAN_ITEM_DELIVERY (DELIVERY_ORDER_ID, PLAN_ITEM_ID, REMARK, PROJECT_ID, CREATE_USER_ID) "
                + "SELECT '" + form.PR_ID + "' as DELIVERY_ORDER_ID, A.PLAN_ITEM_ID as PLAN_ITEM_ID, A.REMARK as REMARK, '" + projectid + "' as PROJECT_ID, '" + createid + "' as CREATE_USER_ID  "
                + "FROM (SELECT pri.PLAN_ITEM_ID, pri.REMARK FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pri.PR_ID = pr.PR_ID " +
                "WHERE pr.PROJECT_ID = '" + projectid + "' AND pr.PR_ID LIKE 'RP%' AND pri.PLAN_ITEM_ID + '/' + pri.REMARK IN (" + ItemId + ") GROUP BY pri.PLAN_ITEM_ID, pri.REMARK)A ";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return form.PR_ID;
            }
        }
        //新增領料數量
        public int refreshDelivery(string deliveryorderid, List<PLAN_ITEM_DELIVERY> lstItem)
        {
            logger.Info("Update delivery items, it's delivery order id =" + deliveryorderid);
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    //將item資料寫入 
                    foreach (PLAN_ITEM_DELIVERY item in lstItem)
                    {
                        PLAN_ITEM_DELIVERY existItem = null;
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("formid", deliveryorderid));
                        parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                        parameters.Add(new SqlParameter("remark", item.REMARK));
                        string sql = "SELECT * FROM PLAN_ITEM_DELIVERY WHERE DELIVERY_ORDER_ID=@formid AND PLAN_ITEM_ID=@itemid AND REMARK =@remark ";
                        logger.Info(sql + " ;" + deliveryorderid + ",plan_item_id=" + item.PLAN_ITEM_ID + ",remark=" + item.REMARK);
                        PLAN_ITEM_DELIVERY excelItem = context.PLAN_ITEM_DELIVERY.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_ITEM_DELIVERY.Find(excelItem.DELIVERY_ID);
                        logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                        existItem.DELIVERY_QTY = item.DELIVERY_QTY;
                        existItem.CREATE_DATE = DateTime.Now;
                        context.PLAN_ITEM_DELIVERY.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update delivery item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new delivery id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return j;
        }

        //取得物料進出紀錄
        public List<PurchaseRequisition> getDeliveryByItemId(string itemid)
        {

            logger.Info(" get receipt record and delivery record by 物料編號 =" + itemid);
            List<PurchaseRequisition> lstItem = new List<PurchaseRequisition>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT A.* FROM (SELECT pid.DELIVERY_ORDER_ID, pid.CREATE_DATE, pid.DELIVERY_QTY, pr.PARENT_PR_ID, ROW_NUMBER() OVER(ORDER BY pid.CREATE_DATE ASC) AS NO " +
                "FROM PLAN_ITEM_DELIVERY pid LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pid.DELIVERY_ORDER_ID = pr.PR_ID WHERE pid.PLAN_ITEM_ID + pid.REMARK =@itemid " +
                "UNION SELECT pri.PR_ID, pr.CREATE_DATE, pri.RECEIPT_QTY, pr.PARENT_PR_ID, ROW_NUMBER() OVER(ORDER BY pr.CREATE_DATE ASC) AS NO FROM PLAN_PURCHASE_REQUISITION_ITEM pri " +
                "LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pri.PR_ID = pr.PR_ID WHERE pr.PR_ID LIKE 'RP%' AND pri.PLAN_ITEM_ID + pri.REMARK =@itemid)A ORDER BY A.CREATE_DATE ";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("itemid", itemid));
            using (var context = new topmepEntities())
            {
                logger.Debug("get receipt and delivery sql=" + sql);
                lstItem = context.Database.SqlQuery<PurchaseRequisition>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get receipt record and delivery record count=" + lstItem.Count);
            return lstItem;
        }

        //寫入領料內容
        public string newDO(string projectid, PLAN_PURCHASE_REQUISITION form, string[] lstItemId)
        {
            //1.建立領料單
            logger.Info("create new delivery form ");
            string sno_key = "DF";
            SerialKeyService snoservice = new SerialKeyService();
            form.PR_ID = snoservice.getSerialKey(sno_key);
            logger.Info("new delivery form =" + form.ToString());
            using (var context = new topmepEntities())
            {
                context.PLAN_PURCHASE_REQUISITION.Add(form);
                int i = context.SaveChanges();
                logger.Debug("Add delivery form=" + i);
                logger.Info("plan delivery form id = " + form.PR_ID);
                //if (i > 0) { status = true; };
                List<topmeperp.Models.PLAN_PURCHASE_REQUISITION_ITEM> lstItem = new List<PLAN_PURCHASE_REQUISITION_ITEM>();
                string ItemId = "";
                for (i = 0; i < lstItemId.Count(); i++)
                {
                    if (i < lstItemId.Count() - 1)
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'" + ",";
                    }
                    else
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'";
                    }
                }

                string sql = "INSERT INTO PLAN_PURCHASE_REQUISITION_ITEM (PR_ID, PLAN_ITEM_ID, RECEIPT_QTY, NEED_DATE, REMARK, ORDER_QTY) "
                + "SELECT '" + form.PR_ID + "' as PR_ID, A.PLAN_ITEM_ID as PLAN_ITEM_ID, A.RECEIPT_QTY as RECEIPT_QTY, A.NEED_DATE as NEED_DATE, A.REMARK as REMARK, A.ORDER_QTY as ORDER_QTY  "
                + "FROM (SELECT pri.* FROM PLAN_PURCHASE_REQUISITION_ITEM pri WHERE pri.PR_ITEM_ID IN (" + ItemId + "))A ";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return form.PR_ID;
            }
        }

        //PLAN_PURCHASE_REQUISITION PRform = null;
        //取得新增的領料單號
        //public PLAN_PURCHASE_REQUISITION getNewDeliveryOrderId(string prid, DateTime createDate)
        //{
        //using (var context = new topmepEntities())
        //{
        //PRform = context.PLAN_PURCHASE_REQUISITION.SqlQuery("SELECT * FROM PLAN_PURCHASE_REQUISITION pr WHERE pr.PARENT_PR_ID =@prid " +
        //"AND pr.PR_ID LIKE 'DF%' AND pr.CREATE_DATE = @createDate "
        //, new SqlParameter("prid", prid), new SqlParameter("createDate", createDate)).FirstOrDefault();
        //}
        //return PRform;
        //}

        //取得領料單資料
        public List<PRFunction> getDOByPrjId(string projectid, string recipient, string prid, string caution)
        {

            logger.Info("search delivery form by 領料說明 =" + caution + ", 領料單編號 =" + prid + ", 領料人所屬單位 =" + recipient);
            List<PRFunction> lstForm = new List<PRFunction>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT CONVERT(char(10), CREATE_DATE, 111) AS CREATE_DATE, PR_ID, RECIPIENT, CAUTION, PROJECT_ID, ROW_NUMBER() OVER(ORDER BY PR_ID) AS NO " +
                "FROM PLAN_PURCHASE_REQUISITION WHERE PROJECT_ID =@projectid AND PR_ID LIKE 'D%' ";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectid));
            //領料說明查詢條件
            if (null != caution && caution != "")
            {
                sql = sql + "AND CAUTION LIKE @caution ";
                parameters.Add(new SqlParameter("caution", '%' + caution + '%'));
            }
            //領料單編號條件
            if (null != prid && prid != "")
            {
                sql = sql + "AND PR_ID =@prid ";
                parameters.Add(new SqlParameter("prid", prid));
            }
            //領料人條件
            if (null != recipient && recipient != "")
            {
                sql = sql + "AND RECIPIENT LIKE @recipient ";
                parameters.Add(new SqlParameter("recipient", '%' + recipient + '%'));
            }
            using (var context = new topmepEntities())
            {
                logger.Debug("get delivery form sql=" + sql);
                lstForm = context.Database.SqlQuery<PRFunction>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get delivery form count=" + lstForm.Count);
            return lstForm;
        }

        //取得個別物料的庫存數量
        public PurchaseRequisition getInventoryByItemId(string itemid)
        {

            logger.Info("search item inventory by planitemid  =" + itemid);
            PurchaseRequisition lstItem = new PurchaseRequisition();
            //處理SQL 預先填入專案代號,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<PurchaseRequisition>("SELECT pri.PLAN_ITEM_ID, pi.ITEM_ID, pi.ITEM_DESC, pi.ITEM_UNIT, pi.SYSTEM_MAIN, SUM(pri.RECEIPT_QTY) - ISNULL(A.DELIVERY_QTY, 0) AS INVENTORY_QTY, SUM(pri.RECEIPT_QTY) AS ALL_RECEIPT_QTY " +
                "FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_ITEM pi ON pri.PLAN_ITEM_ID = pi.PLAN_ITEM_ID " +
                "LEFT JOIN (SELECT pid.PLAN_ITEM_ID, SUM(pid.DELIVERY_QTY) AS DELIVERY_QTY FROM PLAN_ITEM_DELIVERY pid " +
                "GROUP BY pid.PLAN_ITEM_ID)A ON pri.PLAN_ITEM_ID = A.PLAN_ITEM_ID GROUP BY pri.PLAN_ITEM_ID, A.DELIVERY_QTY, " +
                "pi.ITEM_ID, pi.ITEM_DESC, pi.ITEM_UNIT, pi.SYSTEM_MAIN HAVING pri.PLAN_ITEM_ID =@itemid; "
            , new SqlParameter("itemid", itemid)).First();
            }

            return lstItem;
        }

        //取得特定申購單3天內須驗收的物料品項
        public List<PurchaseRequisition> getPlanItemByNeedDate(string prid, string prjid, int status)
        {

            logger.Info(" get materials ready to receive in 3 days by 申購單編號 =" + prid + ",and project id =" + prjid + ",and status = " + status);
            List<PurchaseRequisition> lstItem = new List<PurchaseRequisition>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT pri.NEED_QTY, CONVERT(char(10), pri.NEED_DATE, 111) AS NEED_DATE, pri.REMARK, pri.PR_ID, ISNULL(rp.RECEIPT_QTY, 0) AS RECEIPT_QTY, " +
                    "pi.PLAN_ITEM_ID, pi.ITEM_DESC, pi.ITEM_ID, pi.ITEM_UNIT, pi.SUPPLIER_ID FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_ITEM pi " +
                    "ON pri.PLAN_ITEM_ID = pi.PLAN_ITEM_ID LEFT JOIN (SELECT pri.PLAN_ITEM_ID, pri.REMARK, pri.NEED_QTY, pr.PARENT_PR_ID, main.RECEIPT_QTY " +
                    "FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pri.PR_ID = pr.PR_ID LEFT JOIN " +
                    "(SELECT pri.PLAN_ITEM_ID, pri.REMARK, pr.PARENT_PR_ID, SUM(pri.RECEIPT_QTY)AS RECEIPT_QTY FROM PLAN_PURCHASE_REQUISITION_ITEM pri " +
                    "LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pri.PR_ID = pr.PR_ID WHERE pr.PROJECT_ID =@prjid AND pri.PR_ID LIKE 'RP%' " +
                    "GROUP BY pri.PLAN_ITEM_ID, pri.REMARK, pr.PARENT_PR_ID)main ON main.PLAN_ITEM_ID + main.REMARK + main.PARENT_PR_ID = pri.PLAN_ITEM_ID + pri.REMARK + pri.PR_ID " +
                    "WHERE pr.PROJECT_ID =@prjid AND pri.PR_ID LIKE 'PPO%' GROUP BY pri.PLAN_ITEM_ID, pri.REMARK, pr.PARENT_PR_ID, main.RECEIPT_QTY, pri.NEED_QTY)rp " +
                    "ON pri.PLAN_ITEM_ID + pri.REMARK + pri.PR_ID = rp.PLAN_ITEM_ID + rp.REMARK + rp.PARENT_PR_ID ";
            if (status == 10)
            {
                sql = sql + "WHERE pri.PR_ID =@prid AND pri.NEED_DATE BETWEEN CAST(convert(varchar, getdate()-3, 120) AS DATETIME) AND CAST(convert(varchar, getdate(), 120) AS DATETIME) ";
            }
            else
            {
                sql = sql + "WHERE pri.PR_ID =@prid AND pri.PLAN_ITEM_ID IN (SELECT pri.PLAN_ITEM_ID FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_PURCHASE_REQUISITION pr " +
                 "ON pri.PR_ID = pr.PR_ID WHERE pri.PR_ID LIKE 'PPO%' AND pr.PARENT_PR_ID =@prid) " +
                 "AND pri.NEED_DATE BETWEEN CAST(convert(varchar, getdate()-3, 120) AS DATETIME) AND CAST(convert(varchar, getdate(), 120) AS DATETIME) ";
            }
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("prid", prid));
            parameters.Add(new SqlParameter("prjid", prjid));
            using (var context = new topmepEntities())
            {
                logger.Debug("get materials ready to receive in 3 days sql=" + sql);
                lstItem = context.Database.SqlQuery<PurchaseRequisition>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get material's item count ready to receive in 3 days =" + lstItem.Count);
            return lstItem;
        }

        //取得特定申購單尚未驗收的物料品項
        public List<PurchaseRequisition> getItemNeedReceivedByPrId(string prid, string prjid, int status)
        {

            logger.Info(" get materials need received by 申購單編號 =" + prid + ",and project id =" + prjid + ",and status = " + status);
            List<PurchaseRequisition> lstItem = new List<PurchaseRequisition>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT pri.PLAN_ITEM_ID, pri.REMARK, pri.NEED_QTY, pr.PR_ID, ISNULL(sub.RECEIPT_QTY,0) AS RECEIPT_QTY, pri.NEED_QTY - ISNULL(sub.RECEIPT_QTY, 0) AS diffQty, " +
                 "CONVERT(char(10), pri.NEED_DATE, 111) AS NEED_DATE, pi.ITEM_DESC, pi.ITEM_ID, pi.ITEM_UNIT, pi.SUPPLIER_ID FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_PURCHASE_REQUISITION pr " +
                 "ON pri.PR_ID = pr.PR_ID LEFT JOIN PLAN_ITEM pi ON pri.PLAN_ITEM_ID = pi.PLAN_ITEM_ID LEFT JOIN (SELECT pri.PLAN_ITEM_ID, pri.REMARK, pr.PARENT_PR_ID, main.RECEIPT_QTY " +
                 "FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pri.PR_ID = pr.PR_ID LEFT JOIN " +
                 "(SELECT pri.PLAN_ITEM_ID, pri.REMARK, pr.PARENT_PR_ID, SUM(pri.RECEIPT_QTY)AS RECEIPT_QTY FROM PLAN_PURCHASE_REQUISITION_ITEM pri " +
                 "LEFT JOIN PLAN_PURCHASE_REQUISITION pr ON pri.PR_ID = pr.PR_ID WHERE pr.PROJECT_ID =@prjid AND pri.PR_ID LIKE 'RP%' " +
                 "GROUP BY pri.PLAN_ITEM_ID, pri.REMARK, pr.PARENT_PR_ID)main ON main.PLAN_ITEM_ID + main.REMARK + main.PARENT_PR_ID = pri.PLAN_ITEM_ID + pri.REMARK + pri.PR_ID " +
                 "WHERE pr.PROJECT_ID =@prjid AND pri.PR_ID LIKE 'PPO%')sub ON pri.PLAN_ITEM_ID + pri.REMARK + pr.PR_ID = sub.PLAN_ITEM_ID + sub.REMARK + sub.PARENT_PR_ID ";
            if (status == 10)
            {
                sql = sql + "WHERE pr.PR_ID =@prid AND pri.NEED_QTY - ISNULL(sub.RECEIPT_QTY, 0) > 0 ";
            }
            else
            {
                sql = sql + "WHERE pr.PR_ID =@prid AND pri.PLAN_ITEM_ID IN (SELECT pri.PLAN_ITEM_ID FROM PLAN_PURCHASE_REQUISITION_ITEM pri LEFT JOIN PLAN_PURCHASE_REQUISITION pr " +
                 "ON pri.PR_ID = pr.PR_ID WHERE pri.PR_ID LIKE 'PPO%' AND pr.PARENT_PR_ID =@prid) " +
                 "AND pri.NEED_QTY - ISNULL(sub.RECEIPT_QTY, 0) > 0 ";
            }
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("prid", prid));
            parameters.Add(new SqlParameter("prjid", prjid));
            using (var context = new topmepEntities())
            {
                logger.Debug("get materials need received sql=" + sql);
                lstItem = context.Database.SqlQuery<PurchaseRequisition>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get material's item count ready to receive in 3 days =" + lstItem.Count);
            return lstItem;
        }
        #endregion

        #region 估驗

        public PLAN_ESTIMATION_FORM formEST = null;
        public List<EstimationForm> ESTItem = null;

        string sno_key = "EST";
        public string getEstNo()
        {
            string estNo = null;
            //取得估驗單編號
            using (var context = new topmepEntities())
            {
                SerialKeyService snoservice = new SerialKeyService();
                estNo = snoservice.getSerialKey(sno_key);
            }
            return estNo;
        }

        // 寫入估驗內容
        public string newEST(string formid, PLAN_ESTIMATION_FORM form)//, string[] lstItemId)
        {
            //1.建立估驗單
            logger.Info("create new estimation form ");
            using (var context = new topmepEntities())
            {
                context.PLAN_ESTIMATION_FORM.AddOrUpdate(form);
                int i = context.SaveChanges();
                logger.Debug("Add Purchase Requisition=" + i);
                //if (i > 0) { status = true; };
                /*
                List<topmeperp.Models.PLAN_ESTIMATION_ITEM> lstItem = new List<PLAN_ESTIMATION_ITEM>();
                string ItemId = "";
                for (i = 0; i < lstItemId.Count(); i++)
                {
                    if (i < lstItemId.Count() - 1)
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'" + ",";
                    }
                    else
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'";
                    }
                }

                string sql = "INSERT INTO PLAN_ESTIMATION_ITEM (EST_FORM_ID, PLAN_ITEM_ID) "
                + "SELECT '" + formid + "' as EST_FORM_ID, A.PLAN_ITEM_ID as PLAN_ITEM_ID  "
                + "FROM (SELECT pi.PLAN_ITEM_ID FROM PLAN_ITEM pi WHERE pi.PLAN_ITEM_ID IN (" + ItemId + "))A ";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return formid;
                */
                if (i > 0)
                {
                    return formid;
                }
                else
                {
                    return null;
                }
            }
        }

        //更新估驗數量
        public int refreshEST(string formid, PLAN_ESTIMATION_FORM form, List<PLAN_ESTIMATION_ITEM> lstItem)
        {
            logger.Info("Update plan estimation form id =" + formid);
            int i = 0;
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.Entry(form).State = EntityState.Modified;
                    i = context.SaveChanges();
                    logger.Debug("Update plan estimation form =" + i);
                    logger.Info("purchase estimation item = " + lstItem.Count);
                    //2.將item資料寫入 
                    foreach (PLAN_ESTIMATION_ITEM item in lstItem)
                    {
                        PLAN_ESTIMATION_ITEM existItem = null;
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("formid", formid));
                        parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                        string sql = "SELECT * FROM PLAN_ESTIMATION_ITEM WHERE EST_FORM_ID=@formid AND PLAN_ITEM_ID=@itemid";
                        logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                        PLAN_ESTIMATION_ITEM excelItem = context.PLAN_ESTIMATION_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_ESTIMATION_ITEM.Find(excelItem.EST_ITEM_ID);
                        logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                        existItem.EST_QTY = item.EST_QTY;
                        context.PLAN_ESTIMATION_ITEM.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update plan estimation item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new estimation form id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }

        //取得符合條件之估驗單名單
        public List<ESTFunction> getESTListByEstId(string projectid, string contractid, string estid, int status, string supplier)
        {
            logger.Info("search estimation form by 估驗單編號 =" + estid + ", 合約名稱 =" + contractid + ", 估驗單狀態 =" + status);
            List<ESTFunction> lstForm = new List<ESTFunction>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            if (20 == status)
            {
                string sql = "SELECT CONVERT(char(10), A.CREATE_DATE, 111) AS CREATE_DATE, A.EST_FORM_ID, A.STATUS, A.CONTRACT_NAME, A.SUPPLIER_NAME, ROW_NUMBER() OVER(ORDER BY A.EST_FORM_ID) AS NO " +
                    "FROM (SELECT ef.CREATE_DATE, ef.EST_FORM_ID, ef.STATUS, f.FORM_NAME AS CONTRACT_NAME, f.SUPPLIER_ID AS SUPPLIER_NAME " +
                    "FROM PLAN_ESTIMATION_FORM ef LEFT JOIN PLAN_SUP_INQUIRY f ON ef.CONTRACT_ID = f.INQUIRY_FORM_ID WHERE ef.PROJECT_ID =@projectid)A ";


                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", projectid));
                sql = sql + "WHERE A.STATUS > 10 ";

                //估驗單編號條件
                if (null != estid && estid != "")
                {
                    sql = sql + "AND A.EST_FORM_ID =@estid ";
                    parameters.Add(new SqlParameter("estid", estid));
                }
                //發包項目名稱條件
                if (null != contractid && contractid != "")
                {
                    sql = sql + "AND A.CONTRACT_NAME LIKE @contractid ";
                    parameters.Add(new SqlParameter("contractid", '%' + contractid + '%'));
                }
                //供應商名稱條件
                if (null != supplier && supplier != "")
                {
                    sql = sql + "AND A.SUPPLIER_NAME LIKE @supplier ";
                    parameters.Add(new SqlParameter("supplier", '%' + supplier + '%'));
                }
                using (var context = new topmepEntities())
                {
                    logger.Debug("get estimation form sql=" + sql);
                    lstForm = context.Database.SqlQuery<ESTFunction>(sql, parameters.ToArray()).ToList();
                }
                logger.Info("get estimation form count=" + lstForm.Count);
            }
            else
            {
                string sql = "SELECT CONVERT(char(10), A.CREATE_DATE, 111) AS CREATE_DATE, A.EST_FORM_ID, A.STATUS, A.CONTRACT_NAME, A.SUPPLIER_NAME, ROW_NUMBER() OVER(ORDER BY A.EST_FORM_ID) AS NO " +
                    "FROM (SELECT ef.CREATE_DATE, ef.EST_FORM_ID, ef.STATUS, f.FORM_NAME AS CONTRACT_NAME, f.SUPPLIER_ID AS SUPPLIER_NAME " +
                    "FROM PLAN_ESTIMATION_FORM ef LEFT JOIN PLAN_SUP_INQUIRY f ON ef.CONTRACT_ID = f.INQUIRY_FORM_ID WHERE ef.PROJECT_ID =@projectid)A ";

                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", projectid));
                sql = sql + "WHERE A.STATUS < 20 ";

                using (var context = new topmepEntities())
                {
                    logger.Debug("get estimation form sql=" + sql);
                    lstForm = context.Database.SqlQuery<ESTFunction>(sql, parameters.ToArray()).ToList();
                }
                logger.Info("get estimation form count=" + lstForm.Count);
            }
            return lstForm;
        }

        //取得估驗單資料
        public void getESTByEstId(string estid)
        {
            logger.Info("get form : formid=" + estid);
            formEST = null;
            using (var context = new topmepEntities())
            {
                //取得估驗單檔頭資訊
                string sql = "SELECT * FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID =@estid ";

                formEST = context.PLAN_ESTIMATION_FORM.SqlQuery(sql, new SqlParameter("estid", estid)).FirstOrDefault();
                //取得估驗單明細
                ESTItem = context.Database.SqlQuery<EstimationForm>("SELECT pei.PLAN_ITEM_ID, pei.EST_QTY, pi.ITEM_ID, pi.ITEM_DESC, pi.ITEM_UNIT, psi.ITEM_QTY AS mapQty, pi.ITEM_UNIT_COST, " +
                    "pi.MAN_PRICE, ISNULL(A.CUM_QTY, 0) AS CUM_EST_QTY FROM PLAN_ESTIMATION_ITEM pei LEFT JOIN PLAN_ESTIMATION_FORM ef ON pei.EST_FORM_ID = ef.EST_FORM_ID LEFT JOIN PLAN_ITEM pi ON " +
                    "pei.PLAN_ITEM_ID = pi.PLAN_ITEM_ID LEFT JOIN (SELECT PLAN_ITEM_ID, ITEM_QTY FROM PLAN_SUP_INQUIRY_ITEM WHERE INQUIRY_FORM_ID = '" + formEST.CONTRACT_ID + "')psi ON pei.PLAN_ITEM_ID = psi.PLAN_ITEM_ID " +
                    "LEFT JOIN (SELECT pei.PLAN_ITEM_ID, sum(pei.EST_QTY) AS CUM_QTY FROM PLAN_ESTIMATION_ITEM pei JOIN PLAN_ESTIMATION_FORM ef ON " +
                    "pei.EST_FORM_ID = ef.EST_FORM_ID WHERE ef.CREATE_DATE < (select CREATE_DATE from PLAN_ESTIMATION_FORM where EST_FORM_ID = @estid) GROUP BY  pei.PLAN_ITEM_ID)A " +
                    "ON pei.PLAN_ITEM_ID = A.PLAN_ITEM_ID WHERE pei.EST_FORM_ID = @estid", new SqlParameter("estid", estid)).ToList();
                logger.Debug("get estimation form item count:" + ESTItem.Count);
            }
        }

        public int addOtherPayment(List<PLAN_OTHER_PAYMENT> lstItem)
        {
            //1.新增其他扣款資料
            int i = 0;
            logger.Info("add other payment = " + lstItem.Count);
            //2.將扣款資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_OTHER_PAYMENT item in lstItem)
                {
                    item.TYPE = "O";
                    item.CREATE_DATE = DateTime.Now;
                    context.PLAN_OTHER_PAYMENT.Add(item);
                }
                i = context.SaveChanges();
                //彙整資料至表頭
                string sql = @"
UPDATE PLAN_ESTIMATION_FORM SET
OTHER_PAYMENT=(select SUM(AMOUNT) from PLAN_OTHER_PAYMENT WHERE TYPE='O' AND EST_FORM_ID=@formId)
WHERE EST_FORM_ID=@formId
";
                context.Database.ExecuteSqlCommand(sql, new SqlParameter("formId", lstItem[0].EST_FORM_ID));
                context.SaveChanges();
            }
            logger.Info("add other payment count =" + i);
            return i;
        }
        //取得估驗單其他扣款明細資料
        public List<PLAN_OTHER_PAYMENT> getOtherPayById(string id)
        {

            logger.Info("get other payment by EST id =" + id);
            List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<PLAN_OTHER_PAYMENT>("SELECT * FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID =@id AND TYPE = 'O' ; "
            , new SqlParameter("id", id)).ToList();
            }

            return lstItem;
        }

        //取得估驗單狀態
        public int getStatusById(string id)
        {
            int status = -10;
            logger.Info("get EST status by EST id =" + id);
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                status = context.Database.SqlQuery<int>("SELECT STATUS FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID =@id ; "
            , new SqlParameter("id", id)).FirstOrDefault();
            }

            return status;
        }
        public int delOtherPayByESTId(string estid)
        {
            logger.Info("remove all other payment detail by EST FORM ID=" + estid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete these other payment record by est form id=" + estid);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID=@estid AND TYPE = 'O' ", new SqlParameter("@estid", estid));
            }
            logger.Debug("delete PLAN OTHER PAYMENT count=" + i);
            return i;
        }

        //更新估驗數量
        public int refreshESTQty(string formid, List<PLAN_ESTIMATION_ITEM> lstItem)
        {
            logger.Info("Update estiomation items, it's est form id =" + formid);
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    //將item資料寫入 
                    foreach (PLAN_ESTIMATION_ITEM item in lstItem)
                    {
                        PLAN_ESTIMATION_ITEM existItem = null;
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("formid", formid));
                        parameters.Add(new SqlParameter("itemid", item.PLAN_ITEM_ID));
                        string sql = "SELECT * FROM PLAN_ESTIMATION_ITEM WHERE EST_FORM_ID=@formid AND PLAN_ITEM_ID=@itemid";
                        logger.Info(sql + " ;" + formid + ",plan_item_id=" + item.PLAN_ITEM_ID);
                        PLAN_ESTIMATION_ITEM excelItem = context.PLAN_ESTIMATION_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_ESTIMATION_ITEM.Find(excelItem.EST_ITEM_ID);

                        logger.Debug("find exist item=" + existItem.PLAN_ITEM_ID);
                        if (item.EST_QTY != null)
                        {
                            existItem.EST_QTY = item.EST_QTY;
                        }
                        context.PLAN_ESTIMATION_ITEM.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update estimation item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new est form id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return j;
        }

        //取得新估驗單估驗次數
        int est = 0;
        public int getEstCountById(string contractid)
        {
            using (var context = new topmepEntities())
            {
                est = context.Database.SqlQuery<int>("SELECT ISNULL((SELECT COUNT(CONTRACT_ID) FROM PLAN_ESTIMATION_FORM " +
                    "GROUP BY CONTRACT_ID HAVING CONTRACT_ID =@contractid),0)+1 AS EST_COUNT "
                   , new SqlParameter("contractid", contractid)).First();
            }
            return est;
        }

        //取得現有估驗單之估驗次數
        int estcount = 0;
        public int getEstCountByESTId(string estid)
        {
            using (var context = new topmepEntities())
            {
                estcount = context.Database.SqlQuery<int>("SELECT ISNULL((SELECT COUNT(ef.CONTRACT_ID) + 1 FROM PLAN_ESTIMATION_FORM ef WHERE ef.CREATE_DATE < " +
                    "(select CREATE_DATE from PLAN_ESTIMATION_FORM where EST_FORM_ID =@estid) AND ef.CONTRACT_ID = (select CONTRACT_ID from PLAN_ESTIMATION_FORM " +
                    "where EST_FORM_ID =@estid) GROUP BY ef.CONTRACT_ID),1) AS EST_COUNT  "
                   , new SqlParameter("estid", estid)).First();
            }
            return estcount;
        }

        AdvancePaymentFunction advancePay = null;
        //取得估驗單預付款明細資料
        public AdvancePaymentFunction getAdvancePayById(string id, string contractid)
        {

            logger.Info("get advance payment by EST id + contractid  =" + id);
            using (var context = new topmepEntities())
            {
                advancePay = context.Database.SqlQuery<AdvancePaymentFunction>("SELECT (SELECT AMOUNT FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID =@id AND TYPE = 'A') AS A_AMOUNT, " +
                    "(SELECT AMOUNT FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID =@id AND TYPE = 'B') AS B_AMOUNT, " +
                    "(SELECT AMOUNT FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID =@id AND TYPE = 'C') AS C_AMOUNT, " +
                    "ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE CONTRACT_ID = @contractid AND TYPE = 'A' " +
                    "AND CREATE_DATE < ISNULL((SELECT CREATE_DATE FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID = @id AND TYPE = 'A'), GETDATE()) GROUP BY CONTRACT_ID),0) CUM_A_AMOUNT, " +
                    "ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE CONTRACT_ID = @contractid AND TYPE = 'B' " +
                    "AND CREATE_DATE < ISNULL((SELECT CREATE_DATE FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID = @id AND TYPE = 'B'), GETDATE()) GROUP BY CONTRACT_ID),0) CUM_B_AMOUNT, " +
                    "ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE CONTRACT_ID = @contractid AND TYPE = 'C' " +
                    "AND CREATE_DATE < ISNULL((SELECT CREATE_DATE FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID = @id AND TYPE = 'C'), GETDATE()) GROUP BY CONTRACT_ID),0) CUM_C_AMOUNT  "
            , new SqlParameter("id", id), new SqlParameter("contractid", contractid)).First();
            }

            return advancePay;
        }

        public int addAdvancePayment(List<PLAN_OTHER_PAYMENT> lstItem)
        {
            //1.新增預付款資料
            int i = 0;
            string resetSql = "DELETE PLAN_OTHER_PAYMENT WHERE   EST_FORM_ID=@formId AND CONTRACT_ID=@contractId ";
            logger.Info("add advance payment = " + lstItem.Count);
            //2.將預付款資料寫入 
            using (var context = new topmepEntities())
            {
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("formId", lstItem[0].EST_FORM_ID));
                parameters.Add(new SqlParameter("contractId", lstItem[0].CONTRACT_ID));
                context.Database.ExecuteSqlCommand(resetSql, parameters.ToArray());
                foreach (PLAN_OTHER_PAYMENT item in lstItem)
                {
                    item.CREATE_DATE = DateTime.Now;
                    context.PLAN_OTHER_PAYMENT.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add advance payment count =" + i);
            return i;
        }

        //取得估驗單是否有預付款資料
        public List<PLAN_OTHER_PAYMENT> getAdvancePayByESTId(string id)
        {

            logger.Info("get advance payment by EST id =" + id);
            List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<PLAN_OTHER_PAYMENT>("SELECT * FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID =@id AND TYPE IN ('A','B','C') ; "
            , new SqlParameter("id", id)).ToList();
            }

            return lstItem;
        }

        //更新預付款資料
        public int updateAdvancePayment(string estid, List<PLAN_OTHER_PAYMENT> lstItem)
        {
            //1.修改預付款資料
            int i = 0;
            logger.Info("update advance payment = " + lstItem.Count);
            //2.將預付款資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_OTHER_PAYMENT item in lstItem)
                {
                    PLAN_OTHER_PAYMENT existItem = null;
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("estid", estid));
                    parameters.Add(new SqlParameter("type", item.TYPE));
                    string sql = "SELECT * FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID = @estid and TYPE = @type ";
                    logger.Info(sql + " ;" + item.EST_FORM_ID + item.TYPE);
                    PLAN_OTHER_PAYMENT excelItem = context.PLAN_OTHER_PAYMENT.SqlQuery(sql, parameters.ToArray()).First();
                    existItem = context.PLAN_OTHER_PAYMENT.Find(excelItem.OTHER_PAYMENT_ID);
                    logger.Debug("find exist item=" + existItem.TYPE);
                    existItem.AMOUNT = item.AMOUNT;
                    context.PLAN_OTHER_PAYMENT.AddOrUpdate(existItem);
                }
                i = context.SaveChanges();
            }
            logger.Info("update advance payment count =" + i);
            return i;
        }

        PaymentDetailsFunction detailsPay = null;
        //取得估驗單預付款明細資料
        public PaymentDetailsFunction getDetailsPayById(string formid, string contractid)
        {
            logger.Info("get details payment by  id  =" + formid);
            using (var context = new topmepEntities())
            {
                string sql = @"
                    SELECT D.*,
                    (D.EST_AMOUNT-D.T_FOREIGN+D.T_REPAYMENT) AS SUB_AMOUNT, 
                    (D.CUM_EST_AMOUNT - D.CUM_T_FOREIGN + D.CUM_T_REPAYMENT) AS CUM_SUB_AMOUNT, 
                    (D.EST_AMOUNT-D.T_FOREIGN-D.T_RETENTION-D.T_ADVANCE-D.T_OTHER+D.T_REPAYMENT-D.T_REFUND) AS PAYABLE_AMOUNT, 
                    (D.CUM_EST_AMOUNT-D.CUM_T_FOREIGN-D.CUM_T_RETENTION-D.CUM_T_ADVANCE-D.CUM_T_OTHER+D.CUM_T_REPAYMENT-D.CUM_T_REFUND) AS CUM_PAYABLE_AMOUNT, 
                    (D.EST_AMOUNT-D.T_FOREIGN-D.T_RETENTION-D.T_ADVANCE-D.T_OTHER + D.TAX_AMOUNT+D.T_REPAYMENT-D.T_REFUND) AS PAID_AMOUNT, 
                    (D.CUM_EST_AMOUNT-D.CUM_T_FOREIGN-D.CUM_T_RETENTION-D.CUM_T_ADVANCE-D.CUM_T_OTHER + D.CUM_TAX_AMOUNT+D.CUM_T_REPAYMENT-D.CUM_T_REFUND) AS CUM_PAID_AMOUNT, 
                    (D.T_FOREIGN + D.CUM_T_FOREIGN) AS TOTAL_FOREIGN, (D.EST_AMOUNT - D.T_FOREIGN + D.CUM_EST_AMOUNT - D.CUM_T_FOREIGN + D.T_REPAYMENT + D.CUM_T_REPAYMENT) AS TOTAL_SUB_AMOUNT, 
                    (D.T_RETENTION + D.CUM_T_RETENTION) AS TOTAL_RETENTION, 
                    (D.T_ADVANCE + D.CUM_T_ADVANCE) AS TOTAL_ADVANCE, 
                    (D.T_REPAYMENT + D.CUM_T_REPAYMENT) AS TOTAL_REPAYMENT, 
                    (D.T_REFUND + D.CUM_T_REFUND) AS TOTAL_REFUND, 
                    (D.T_OTHER + D.CUM_T_OTHER) AS TOTAL_OTHER,
                    (D.EST_AMOUNT-D.T_FOREIGN-D.T_RETENTION-D.T_ADVANCE-D.T_OTHER + D.CUM_EST_AMOUNT-D.CUM_T_FOREIGN - D.CUM_T_RETENTION - D.CUM_T_ADVANCE - D.CUM_T_OTHER + D.CUM_T_REPAYMENT - D.CUM_T_REFUND) AS TOTAL_PAYABLE_AMOUNT,  
                    (D.TAX_AMOUNT + D.CUM_TAX_AMOUNT) AS TOTAL_TAX_AMOUNT, 
                    (D.EST_AMOUNT-D.T_FOREIGN-D.T_RETENTION-D.T_ADVANCE-D.T_OTHER + D.TAX_AMOUNT + D.CUM_EST_AMOUNT-D.CUM_T_FOREIGN-D.CUM_T_RETENTION-D.CUM_T_ADVANCE-D.CUM_T_OTHER + D.CUM_TAX_AMOUNT + D.CUM_T_REPAYMENT - D.CUM_T_REFUND) AS TOTAL_PAID_AMOUNT 
                    FROM(SELECT ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT 
                    WHERE EST_FORM_ID =@formid AND CONTRACT_ID = @contractid AND TYPE IN('A', 'B', 'C') GROUP BY EST_FORM_ID, CONTRACT_ID), 0) AS T_ADVANCE,
                    ISNULL((SELECT SUM(pop.AMOUNT) FROM PLAN_OTHER_PAYMENT pop LEFT JOIN PLAN_ESTIMATION_FORM f ON pop.EST_FORM_ID = f.EST_FORM_ID WHERE pop.CONTRACT_ID = @contractid AND pop.TYPE IN('A', 'B', 'C') AND f.STATUS >= 0 
                    AND pop.CREATE_DATE < (SELECT ISNULL((SELECT MIN(CREATE_DATE) FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID = @formid AND TYPE IN('A', 'B', 'C')), GETDATE())) GROUP BY pop.CONTRACT_ID),0) AS CUM_T_ADVANCE, 
                    ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID =@formid AND CONTRACT_ID = @contractid AND TYPE = 'R' GROUP BY EST_FORM_ID, CONTRACT_ID), 0) AS T_REPAYMENT, 
                    ISNULL((SELECT SUM(pop.AMOUNT) FROM PLAN_OTHER_PAYMENT pop LEFT JOIN PLAN_ESTIMATION_FORM f ON pop.EST_FORM_ID = f.EST_FORM_ID WHERE pop.CONTRACT_ID = @contractid AND pop.TYPE = 'R' AND f.STATUS >= 0 
                    AND pop.CREATE_DATE < (SELECT ISNULL((SELECT MIN(CREATE_DATE) FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID = @formid AND TYPE = 'R'), GETDATE())) GROUP BY pop.CONTRACT_ID),0) AS CUM_T_REPAYMENT, 
                    ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID =@formid AND CONTRACT_ID = @contractid AND TYPE = 'F' GROUP BY EST_FORM_ID, CONTRACT_ID), 0) AS T_REFUND, 
                    ISNULL((SELECT SUM(pop.AMOUNT) FROM PLAN_OTHER_PAYMENT pop LEFT JOIN PLAN_ESTIMATION_FORM f ON pop.EST_FORM_ID = f.EST_FORM_ID WHERE pop.CONTRACT_ID = @contractid AND pop.TYPE = 'F' AND f.STATUS >= 0 
                    AND pop.CREATE_DATE < (SELECT ISNULL((SELECT MIN(CREATE_DATE) FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID = @formid AND TYPE = 'F'), GETDATE())) GROUP BY pop.CONTRACT_ID),0) AS CUM_T_REFUND, 
                    ISNULL((SELECT SUM(AMOUNT) AS AMOUNT FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID IN (SELECT EST_FORM_ID FROM PLAN_ESTIMATION_FORM WHERE CREATE_DATE < (SELECT DISTINCT ef.CREATE_DATE FROM PLAN_OTHER_PAYMENT pop JOIN PLAN_ESTIMATION_FORM ef 
                    ON pop.EST_FORM_ID = ef.EST_FORM_ID WHERE pop.EST_FORM_ID = @formid AND pop.TYPE = 'O' AND pop.CONTRACT_ID = @contractid) AND STATUS >= 0)),0) AS CUM_T_OTHER, 
                    ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE TYPE = 'O' GROUP BY EST_FORM_ID HAVING EST_FORM_ID = @formid),0) AS T_OTHER, 
                    ISNULL((SELECT TAX_RATIO FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = @formid),0) AS TAX_RATIO, 
                    ISNULL((SELECT RETENTION_PAYMENT FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = @formid),0) AS T_RETENTION, 
                    ISNULL((SELECT TAX_AMOUNT FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = @formid),0) AS TAX_AMOUNT, 
                    ISNULL((SELECT IIF(loan.AMOUNT < 0 , -loan.AMOUNT, 0) FROM PLAN_ESTIMATION_FORM f LEFT JOIN (SELECT l.BANK_NAME, SUM(TRANSACTION_TYPE*AMOUNT) AS AMOUNT FROM FIN_LOAN_TRANACTION t LEFT JOIN FIN_BANK_LOAN l ON t.BL_ID = l.BL_ID 
                    WHERE ISNULL(l.IS_SUPPLIER, 'N') = 'Y' GROUP BY t.BL_ID, l.BANK_NAME)loan ON f.PAYEE = loan.BANK_NAME WHERE f.EST_FORM_ID =@formid),0) AS LOAN_AMOUNT, 
                    ISNULL((SELECT loan.BL_ID FROM PLAN_ESTIMATION_FORM f LEFT JOIN (SELECT l.BANK_NAME, t.BL_ID, SUM(TRANSACTION_TYPE*AMOUNT) AS AMOUNT FROM FIN_LOAN_TRANACTION t LEFT JOIN FIN_BANK_LOAN l ON t.BL_ID = l.BL_ID 
                    WHERE ISNULL(l.IS_SUPPLIER, 'N') = 'Y' GROUP BY t.BL_ID, l.BANK_NAME)loan ON f.PAYEE = loan.BANK_NAME WHERE f.EST_FORM_ID =@formid),0) AS LOAN_PAYEE_ID, 
                    ISNULL((SELECT SUM(RETENTION_PAYMENT) FROM PLAN_ESTIMATION_FORM WHERE CONTRACT_ID = @contractid AND STATUS >= 0 AND CREATE_DATE < (SELECT CREATE_DATE FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = @formid)),0)  AS CUM_T_RETENTION, 
                    ISNULL((SELECT SUM(TAX_AMOUNT) FROM PLAN_ESTIMATION_FORM WHERE CONTRACT_ID = @contractid AND STATUS >= 0 AND CREATE_DATE < (SELECT CREATE_DATE FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = @formid)),0) AS CUM_TAX_AMOUNT, 
                    ISNULL((SELECT ISNULL(PAYMENT_TRANSFER, 0) + ISNULL(FOREIGN_PAYMENT, 0) - ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID =@formid AND CONTRACT_ID =@contractid AND TYPE = 'R' GROUP BY EST_FORM_ID, CONTRACT_ID), 0) PRICE 
                    FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = @formid),0) AS EST_AMOUNT, 
                    ISNULL((SELECT SUM(PAYMENT_TRANSFER) + SUM(FOREIGN_PAYMENT) - ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID =@formid AND CONTRACT_ID =@contractid AND TYPE = 'R' GROUP BY EST_FORM_ID, CONTRACT_ID), 0) PRICE 
                    FROM PLAN_ESTIMATION_FORM WHERE STATUS >= 0 AND CREATE_DATE < (SELECT CREATE_DATE FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = @formid AND CONTRACT_ID = @contractid)), 0) AS CUM_EST_AMOUNT,  
                    ISNULL((SELECT FOREIGN_PAYMENT FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = @formid),0) AS T_FOREIGN, ISNULL((SELECT SUM(FOREIGN_PAYMENT) FROM PLAN_ESTIMATION_FORM WHERE CONTRACT_ID = @contractid AND STATUS >= 0 
                    AND CREATE_DATE < (SELECT CREATE_DATE FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = @formid)),0) AS CUM_T_FOREIGN)D 
                    ";
                logger.Debug("sql=" + sql);
                detailsPay = context.Database.SqlQuery<PaymentDetailsFunction>(sql, new SqlParameter("formid", formid), new SqlParameter("contractid", contractid)).First();
            }
            return detailsPay;
        }

        decimal retention = 0;
        public decimal getRetentionAmountById(string id)
        {
            using (var context = new topmepEntities())
            {
                retention = context.Database.SqlQuery<decimal>("SELECT RATIO * AMOUNT / 100 * 1.05 FROM(SELECT ISNULL((SELECT PAYMENT_RETENTION_RATIO FROM PLAN_PAYMENT_TERMS WHERE CONTRACT_ID = SUBSTRING(@id, 10, LEN(@id) - 9)), 0) AS RATIO, " +
                    "(SELECT ISNULL(SUM(pei.EST_QTY * pi.ITEM_UNIT_COST),0) PRICE FROM PLAN_ESTIMATION_ITEM pei LEFT JOIN PLAN_ITEM pi ON pei.PLAN_ITEM_ID = pi.PLAN_ITEM_ID WHERE pei.EST_FORM_ID = SUBSTRING(@id, 1, 9)) AS AMOUNT)B  "
                   , new SqlParameter("id", id)).FirstOrDefault();
            }
            return retention;
        }
        //寫入估驗保留款
        public int UpdateRetentionAmountById(string formid, string contractid)
        {
            int i = 0;
            logger.Info("update retention payment of EST form by id" + formid);
            string sql = "UPDATE  PLAN_ESTIMATION_FORM SET RETENTION_PAYMENT = i.PAY, TAX_AMOUNT = i.TAX_AMOUNT  " +
                "FROM(SELECT ROUND(CAST(RATIO * AMOUNT / 100 * (1+TAX_RATIO/100) AS decimal(10,1)),0) AS PAY, ROUND(CAST((AMOUNT -T_FOREIGN) * TAX_RATIO/100 AS decimal(8,1)),0) AS TAX_AMOUNT FROM(SELECT ISNULL((SELECT PAYMENT_RETENTION_RATIO FROM PLAN_PAYMENT_TERMS WHERE " +
                "CONTRACT_ID = @contractid), 0) AS RATIO, ISNULL((SELECT TAX_RATIO FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = @formid),0) AS TAX_RATIO, " +
                "ISNULL((SELECT FOREIGN_PAYMENT FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = @formid),0) AS T_FOREIGN, " +
                "(SELECT ISNULL(PAYMENT_TRANSFER, 0) + ISNULL(FOREIGN_PAYMENT, 0) - ISNULL((SELECT SUM(AMOUNT) FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID =@formid AND CONTRACT_ID =@contractid AND TYPE = 'R' GROUP BY EST_FORM_ID, CONTRACT_ID), 0) PRICE FROM PLAN_ESTIMATION_FORM " +
                "WHERE EST_FORM_ID = @formid) AS AMOUNT)B) i WHERE EST_FORM_ID = @formid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formid", formid));
            parameters.Add(new SqlParameter("contractid", contractid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return i;
        }
        //寫入當次估驗小計金額與實付金額
        public int UpdatePaidAmountById(string formid, decimal paidAmt)
        {
            int i = 0;
            logger.Info("update sub amount and amount need to pay from EST form by id" + formid);
            string sql = "UPDATE PLAN_ESTIMATION_FORM SET PAID_AMOUNT = @paidAmt WHERE EST_FORM_ID = @formid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formid", formid));
            parameters.Add(new SqlParameter("paidAmt", paidAmt));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return i;
        }
        //取得估驗單之狀態
        int status = 0;
        public int getEstStatusByESTId(string estid)
        {
            using (var context = new topmepEntities())
            {
                estcount = context.Database.SqlQuery<int>("SELECT STATUS FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = @estid "
                   , new SqlParameter("estid", estid)).First();
            }
            return status;
        }

        public int delESTByESTId(string estid)
        {
            logger.Info("remove EST form detail by EST FORM ID =" + estid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete EST form record by est form id =" + estid);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_ESTIMATION_FORM WHERE EST_FORM_ID = @estid ", new SqlParameter("@estid", estid));
            }
            logger.Debug("delete PLAN ESTIMATION FORM count=" + i);
            return i;
        }

        public int delESTItemsByESTId(string estid)
        {
            logger.Info("remove EST items detail by EST FORM ID  =" + estid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete EST items record by est form id =" + estid);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_ESTIMATION_ITEM WHERE EST_FORM_ID = @estid ", new SqlParameter("@estid", estid));
            }
            logger.Debug("delete PLAN ESTIMATION ITEM count=" + i);
            return i;
        }
        //更新估驗單狀態為草稿
        public int UpdateESTStatusById(string estid)
        {
            int i = 0;
            logger.Info("update the status of EST form by estid" + estid);
            string sql = "UPDATE  PLAN_ESTIMATION_FORM SET STATUS = 10 WHERE EST_FORM_ID = @estid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("estid", estid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }



        //修改估驗單內容
        public int RefreshESTByEstId(string estid, string tax, decimal taxratio)
        {
            int i = 0;
            logger.Info("update EST form by estid" + estid);
            string sql = "UPDATE  PLAN_ESTIMATION_FORM SET PLUS_TAX = @tax, TAX_RATIO =@taxratio WHERE EST_FORM_ID = @estid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("estid", estid));
            parameters.Add(new SqlParameter("tax", tax));
            parameters.Add(new SqlParameter("taxratio", taxratio));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return i;
        }
        //修改估驗單額外扣款
        //public int RefreshESTAmountByEstId(string estid, decimal subAmount, decimal foreign_payment, decimal retention, decimal tax_amount, string remark)
        public int RefreshESTAmountByEstId(string estid, decimal subAmount, string payee, string projectName, DateTime? paymentDate, string remark, string indirectCostType)
        {
            int i = 0;
            logger.Info("update EST form by estid" + estid);
            string sql = "UPDATE  PLAN_ESTIMATION_FORM SET PAYMENT_TRANSFER = @subAmount, PAYEE = @payee, PROJECT_NAME = @projectName, PAYMENT_DATE =@paymentDate, REMARK =@remark, INDIRECT_COST_TYPE =@indirectCostType WHERE EST_FORM_ID = @estid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("estid", estid));
            parameters.Add(new SqlParameter("subAmount", subAmount));
            parameters.Add(new SqlParameter("payee", payee));
            parameters.Add(new SqlParameter("projectName", projectName));
            parameters.Add(new SqlParameter("paymentDate", paymentDate));
            parameters.Add(new SqlParameter("remark", remark));
            parameters.Add(new SqlParameter("indirectCostType", indirectCostType));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }
        /*
        //更新估驗單狀態為送審
        public int RefreshESTStatusById(string estid)
        {
            int i = 0;
            logger.Info("update the status of EST form by estid" + estid);
            string sql = "UPDATE  PLAN_ESTIMATION_FORM SET STATUS = 20 WHERE EST_FORM_ID = @estid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("estid", estid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }
        //更新估驗單狀態為退件
        public int RejectESTByEstId(string estid)
        {
            int i = 0;
            logger.Info("reject EST form by estid" + estid);
            string sql = "UPDATE  PLAN_ESTIMATION_FORM SET STATUS = 0 WHERE EST_FORM_ID = @estid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("estid", estid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }

        //更新估驗單狀態為已核可
        public int ApproveESTByEstId(string estid)
        {
            int i = 0;
            logger.Info("Approve EST form by estid" + estid);
            string sql = "UPDATE  PLAN_ESTIMATION_FORM SET STATUS = 30, MODIFY_DATE = GETDATE() WHERE EST_FORM_ID = @estid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("estid", estid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }
        */
        //取得估驗單憑證資料
        public List<PLAN_INVOICE> getInvoiceById(string id)
        {

            logger.Info("get invoice by EST id =" + id);
            List<PLAN_INVOICE> lstItem = new List<PLAN_INVOICE>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<PLAN_INVOICE>("SELECT INVOICE_ID, EST_FORM_ID, CONTRACT_ID, INVOICE_DATE, INVOICE_NUMBER, " +
                    "AMOUNT, TAX, TYPE, SUB_TYPE, PLAN_ITEM_ID, DISCOUNT_QTY, DISCOUNT_UNIT_PRICE, CREATE_DATE FROM PLAN_INVOICE WHERE EST_FORM_ID =@id ; "
            , new SqlParameter("id", id)).ToList();
            }

            return lstItem;
        }
        //取得廠商合約憑證資料
        public List<PLAN_INVOICE> getInvoiceByContractId(string formid, string contractid)
        {

            logger.Info("get invoice by contract id =" + contractid);
            List<PLAN_INVOICE> lstItem = new List<PLAN_INVOICE>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<PLAN_INVOICE>("SELECT * FROM PLAN_INVOICE WHERE CONTRACT_ID =@contractid AND CREATE_DATE < (SELECT MIN(CREATE_DATE) FROM PLAN_INVOICE WHERE EST_FORM_ID =@formid) ; "
            , new SqlParameter("contractid", contractid), new SqlParameter("formid", formid)).ToList();
            }

            return lstItem;
        }
        public int addInvoice(List<PLAN_INVOICE> lstItem)
        {
            //1.新增憑證資料
            int i = 0;
            logger.Info("add invoice = " + lstItem.Count);
            //2.將扣款資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_INVOICE item in lstItem)
                {
                    item.CREATE_DATE = DateTime.Now;
                    context.PLAN_INVOICE.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add invoice count =" + i);
            return i;
        }

        //取得特定估驗單憑證張數
        public int getInvoicePiecesById(string estid)
        {
            int pieces = 0;
            logger.Info("get invoice pieces by form id  =" + estid);
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                pieces = context.Database.SqlQuery<int>("SELECT COUNT(*) AS invoicePieces FROM PLAN_INVOICE WHERE EST_FORM_ID =@estid GROUP BY EST_FORM_ID ; "
            , new SqlParameter("estid", estid)).FirstOrDefault();
            }

            return pieces;
        }
        //取得付款條件
        public string getTermsByContractId(string contractid)
        {
            string terms = null;
            logger.Info("get payment terms by contractid  =" + contractid);
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                terms = context.Database.SqlQuery<string>("SELECT PAYMENT_TERMS FROM PLAN_PAYMENT_TERMS WHERE CONTRACT_ID =@contractid ; "
            , new SqlParameter("contractid", contractid)).FirstOrDefault();
            }

            return terms;
        }

        //取得估驗單代付支出明細資料
        public List<RePaymentFunction> getRePaymentById(string id)
        {
            string sql = @"
SELECT pop.AMOUNT AS AMOUNT, pop.REASON AS REASON, pop.CONTRACT_ID_FOR_REFUND AS CONTRACT_ID_FOR_REFUND, s.SUPPLIER_ID AS COMPANY_NAME 
FROM PLAN_OTHER_PAYMENT pop LEFT JOIN PLAN_SUP_INQUIRY s ON pop.CONTRACT_ID_FOR_REFUND = s.INQUIRY_FORM_ID WHERE pop.EST_FORM_ID =@id AND pop.TYPE = 'R' 
";
            logger.Info("get repayment :" + sql + " by EST id =" + id);
            List<RePaymentFunction> lstItem = new List<RePaymentFunction>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<RePaymentFunction>(sql, new SqlParameter("id", id)).ToList();
            }

            return lstItem;
        }

        //取得專案發包廠商資料
        public List<RePaymentFunction> getSupplierOfContractByPrjId(string prjid)
        {

            logger.Info("get repayment by projectid  =" + prjid);
            List<RePaymentFunction> lstItem = new List<RePaymentFunction>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<RePaymentFunction>("SELECT DISTINCT pi.SUPPLIER_ID AS COMPANY_NAME, pi.INQUIRY_FORM_ID AS CONTRACT_NAME " +
                    "FROM PLAN_ITEM pi WHERE PROJECT_ID =@prjid AND SUPPLIER_ID IS NOT NULL UNION SELECT DISTINCT pi.MAN_SUPPLIER_ID, " +
                    "pi.MAN_FORM_ID FROM PLAN_ITEM pi WHERE PROJECT_ID =@prjid AND MAN_SUPPLIER_ID IS NOT NULL ; "
            , new SqlParameter("prjid", prjid)).ToList();
            }

            return lstItem;
        }

        public int AddRePay(List<PLAN_OTHER_PAYMENT> lstItem)
        {
            //1.新增代付支出資料
            int i = 0;
            logger.Info("add repayment = " + lstItem.Count);
            //2.將代付支出資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_OTHER_PAYMENT item in lstItem)
                {
                    item.TYPE = "R";
                    item.CREATE_DATE = DateTime.Now;
                    context.PLAN_OTHER_PAYMENT.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add repayment count =" + i);
            return i;
        }

        public int delRePayByESTId(string estid)
        {
            logger.Info("remove all repayment detail by EST FORM ID=" + estid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete these repayment record by est form id=" + estid);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID=@estid AND TYPE = 'R' ", new SqlParameter("@estid", estid));
            }
            logger.Debug("delete PLAN OTHER PAYMENT count=" + i);
            return i;
        }

        //取得估驗單代付扣回明細資料
        public List<RePaymentFunction> getRefundById(string id)
        {

            logger.Info("get refund by EST id =" + id);
            List<RePaymentFunction> lstItem = new List<RePaymentFunction>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<RePaymentFunction>("SELECT A.*, ISNULL((SELECT COUNT(ef.CONTRACT_ID) + 1 FROM PLAN_ESTIMATION_FORM ef WHERE ef.CREATE_DATE <  " +
                    "(SELECT DISTINCT pef.CREATE_DATE from PLAN_OTHER_PAYMENT pop LEFT JOIN PLAN_ESTIMATION_FORM pef ON pef.EST_FORM_ID = pop.EST_FORM_ID_REFUND WHERE pop.EST_FORM_ID = @id AND pef.CREATE_DATE IS NOT NULL) " +
                    "AND ef.CONTRACT_ID = (SELECT DISTINCT pef.CONTRACT_ID FROM PLAN_OTHER_PAYMENT pop LEFT JOIN PLAN_ESTIMATION_FORM pef ON pef.EST_FORM_ID = pop.EST_FORM_ID_REFUND WHERE pop.EST_FORM_ID = @id AND pef.CREATE_DATE IS NOT NULL) " +
                    "GROUP BY ef.CONTRACT_ID),1) AS EST_COUNT_REFUND FROM(SELECT pop.AMOUNT AS AMOUNT, pop.REASON AS REASON, pop.CONTRACT_ID AS CONTRACT_ID, " +
                    "pop.EST_FORM_ID_REFUND AS EST_FORM_ID_REFUND, pop.CONTRACT_ID_FOR_REFUND AS CONTRACT_ID_FOR_REFUND, s.SUPPLIER_ID AS COMPANY_NAME FROM PLAN_OTHER_PAYMENT pop LEFT JOIN PLAN_SUP_INQUIRY s " +
                    "ON pop.CONTRACT_ID_FOR_REFUND = s.INQUIRY_FORM_ID " +
                    "WHERE pop.EST_FORM_ID =@id AND pop.TYPE = 'F')A ; "
            , new SqlParameter("id", id)).ToList();
            }

            return lstItem;
        }

        //取得特定廠商所有代付扣回明細資料
        public List<RePaymentFunction> getRefundOfSupplierById(string contractid)
        {

            logger.Info("get all refunds of this supplier by contractid =" + contractid);
            List<RePaymentFunction> lstItem = new List<RePaymentFunction>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<RePaymentFunction>("SELECT pop.AMOUNT AS AMOUNT, pop.REASON AS REASON, pop.CONTRACT_ID AS CONTRACT_ID, " +
                    "pop.EST_FORM_ID_REFUND AS EST_FORM_ID_REFUND, pop.EST_COUNT_REFUND AS EST_COUNT_REFUND, s.SUPPLIER_ID AS COMPANY_NAME " +
                    "FROM PLAN_OTHER_PAYMENT pop LEFT JOIN PLAN_SUP_INQUIRY s ON pop.CONTRACT_ID_FOR_REFUND = s.INQUIRY_FORM_ID " +
                    "WHERE pop.CONTRACT_ID = @contractid AND pop.TYPE = 'F' AND pop.EST_COUNT_REFUND IS NOT NULL ; "
            , new SqlParameter("contractid", contractid)).ToList();
            }

            return lstItem;
        }
        //取得需扣回給代付廠商之資料
        public List<RePaymentFunction> getSupplierOfContractRefundById(string contractid)
        {

            logger.Info("get refund by contractid  =" + contractid);
            List<RePaymentFunction> lstItem = new List<RePaymentFunction>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<RePaymentFunction>("SELECT pop.AMOUNT AS AMOUNT, pop.REASON AS REASON, pop.CONTRACT_ID AS CONTRACT_ID, pop.EST_FORM_ID AS EST_FORM_ID_REFUND, " +
                    "pop.OTHER_PAYMENT_ID AS OTHER_PAYMENT_ID, s.SUPPLIER_ID AS COMPANY_NAME FROM PLAN_OTHER_PAYMENT pop LEFT JOIN PLAN_SUP_INQUIRY s " +
                    "ON pop.CONTRACT_ID = s.INQUIRY_FORM_ID " +
                    "WHERE pop.CONTRACT_ID_FOR_REFUND =@contractid AND pop.TYPE = 'R' AND pop.CONTRACT_ID + pop.EST_FORM_ID + pop.CONTRACT_ID_FOR_REFUND + pop.REASON NOT IN " +
                    "(SELECT op.CONTRACT_ID_FOR_REFUND + op.EST_FORM_ID_REFUND + op.CONTRACT_ID + op.REASON FROM PLAN_OTHER_PAYMENT op WHERE op.TYPE = 'F'); "
            , new SqlParameter("contractid", contractid)).ToList();
            }

            return lstItem;
        }

        public string AddRefund(string formid, string contractid, string[] lstItemId)
        {
            //寫入代付支出資料
            using (var context = new topmepEntities())
            {
                List<PLAN_OTHER_PAYMENT> lstItem = new List<PLAN_OTHER_PAYMENT>();
                string ItemId = "";
                for (i = 0; i < lstItemId.Count(); i++)
                {
                    if (i < lstItemId.Count() - 1)
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'" + ",";
                    }
                    else
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'";
                    }
                }

                string sql = "INSERT INTO PLAN_OTHER_PAYMENT (EST_FORM_ID, CONTRACT_ID, AMOUNT, "
                    + "CONTRACT_ID_FOR_REFUND, REASON, EST_FORM_ID_REFUND, TYPE) "
                    + "SELECT '" + formid + "' AS EST_FORM_ID, '" + contractid + "' AS CONTRACT_ID, AMOUNT, CONTRACT_ID_FOR_REFUND, REASON, EST_FORM_ID AS EST_FORM_ID_REFUND, 'F' AS TYPE "
                    + "FROM PLAN_OTHER_PAYMENT where OTHER_PAYMENT_ID IN (" + ItemId + ")";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return formid;
            }
        }

        public int delRefundByESTId(string estid)
        {
            logger.Info("remove all refund detail by EST FORM ID=" + estid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete these refund record by est form id=" + estid);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_OTHER_PAYMENT WHERE EST_FORM_ID=@estid AND TYPE = 'F' ", new SqlParameter("@estid", estid));
            }
            logger.Debug("delete PLAN OTHER PAYMENT count=" + i);
            return i;
        }

        public int RefreshRefund(List<PLAN_OTHER_PAYMENT> lstItem)
        {
            //1.新增代付扣回資料
            int i = 0;
            logger.Info("add refund = " + lstItem.Count);
            //2.將代付扣回資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_OTHER_PAYMENT item in lstItem)
                {
                    item.TYPE = "F";
                    item.CREATE_DATE = DateTime.Now;
                    context.PLAN_OTHER_PAYMENT.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add refund count =" + i);
            return i;
        }

        //取得廠商合約需扣回之總金額
        public decimal getBalanceOfRefundById(string id)
        {

            logger.Info("get the balance of refund by contractid  =" + id);
            decimal balance = 0;
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                balance = context.Database.SqlQuery<decimal>("SELECT ISNULL(R_AMOUNT,0) - ISNULL(F_AMOUNT,0) AS BALANCE " +
                    "FROM (SELECT CONTRACT_ID_FOR_REFUND, SUM(AMOUNT) AS R_AMOUNT FROM PLAN_OTHER_PAYMENT WHERE TYPE = 'R' GROUP BY CONTRACT_ID_FOR_REFUND) B " +
                    "LEFT JOIN (SELECT CONTRACT_ID, SUM(AMOUNT) AS F_AMOUNT FROM PLAN_OTHER_PAYMENT WHERE TYPE = 'F' GROUP BY CONTRACT_ID)A " +
                    "ON REPLACE(B.CONTRACT_ID_FOR_REFUND,'*',',') = A.CONTRACT_ID  GROUP BY B.CONTRACT_ID_FOR_REFUND, R_AMOUNT, F_AMOUNT " +
                    "HAVING REPLACE(B.CONTRACT_ID_FOR_REFUND,'*',',') =@id AND ISNULL(R_AMOUNT,0) - ISNULL(F_AMOUNT,0) > 0 ; "
                   , new SqlParameter("id", id)).FirstOrDefault();
            }

            return balance;
        }

        //取得未核准之估驗單號碼
        public string getEstNoByContractId(string contractid)
        {
            string EstNo = null;
            logger.Info("get payment terms by contractid  =" + contractid);
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                EstNo = context.Database.SqlQuery<string>("SELECT EST_FORM_ID FROM PLAN_ESTIMATION_FORM WHERE CONTRACT_ID =@contractid AND STATUS < 30 AND STATUS >= 0 ; "
            , new SqlParameter("contractid", contractid)).FirstOrDefault();
            }

            return EstNo;
        }

        public int delInvoiceByESTId(string estid)
        {
            logger.Info("remove all invoice detail by EST FORM ID=" + estid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete these invoice record by est form id=" + estid);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_INVOICE WHERE EST_FORM_ID=@estid ", new SqlParameter("@estid", estid));
            }
            logger.Debug("delete PLAN INVOICE count=" + i);
            return i;
        }
        #endregion

        //取得6個月內現金流資料(6個月後資料總結至最大的日期呈現,日期遇假日(即星期六與星期日)會遞延)
        public List<CashFlowFunction> getCashFlow()
        {

            logger.Info("get cash flow order by date in six months !!");
            List<CashFlowFunction> lstItem = new List<CashFlowFunction>();
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                string sql = @"  
    
--DECLARE @EventDate varchar(10)
--set  @EventDate='2018/09/05';

WITH CompanyExpense  AS (
  SELECT dbo.getWorkDay(v.EVENT_DATE) as EVENT_DATE,SUM(AMOUNT) AMOUNT,SUM(AMOUNT_REAL) AMOUNT_REAL 
  FROM (
   SELECT BUDGET.EVENT_DATE,ISNULL(BUDGET.AMOUNT,0) AMOUNT,ISNULL(REALAMT.AMOUNT_REAL,0) AMOUNT_REAL FROM (
	 SELECT 
	  CONVERT (Datetime,CONVERT(varchar, r.CURRENT_YEAR) + '/' + CONVERT(varchar, r.BUDGET_MONTH)
	   + '/' + CONVERT(varchar, S.BUDGET_DAY)) EVENT_DATE,SUBJECT_ID, CURRENT_YEAR AS CURRENT_YEAR, BUDGET_MONTH, AMOUNT 
	 FROM FIN_EXPENSE_BUDGET R JOIN FIN_SUBJECT S 
	 ON R.SUBJECT_ID=S.FIN_SUBJECT_ID WHERE S.BUDGET_DAY IS NOT NULL 
	 UNION         
	 SELECT 
	  	 Dateadd(day,-1,Dateadd(month,1,CONVERT (Datetime,CONVERT(varchar, r.CURRENT_YEAR) + '/' 
		 + CONVERT(varchar, r.BUDGET_MONTH) + '/' + CONVERT(varchar, 1))))  EVENT_DATE,
		 SUBJECT_ID, CURRENT_YEAR AS CURRENT_YEAR, BUDGET_MONTH, AMOUNT 
	 FROM FIN_EXPENSE_BUDGET R  JOIN FIN_SUBJECT S 
	 ON R.SUBJECT_ID=S.FIN_SUBJECT_ID WHERE S.BUDGET_DAY IS NULL 
     ) BUDGET LEFT JOIN (
	  --公司費用
	  SELECT  PAYMENT_DATE EVENT_DATE,i.FIN_SUBJECT_ID SUBJECT_ID, sum(i.AMOUNT) as AMOUNT_REAL 
	  FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM i 
	  ON f.EXP_FORM_ID = i.EXP_FORM_ID WHERE i.FIN_SUBJECT_ID Like '6%'
	  AND f.STATUS = 30 GROUP BY PAYMENT_DATE,FIN_SUBJECT_ID
	  ) REALAMT ON BUDGET.EVENT_DATE=REALAMT.EVENT_DATE AND BUDGET.SUBJECT_ID=REALAMT.SUBJECT_ID
	) V GROUP BY dbo.getWorkDay(v.EVENT_DATE)
 ),
 SiteExpense AS (
  SELECT dbo.getWorkDay(v.EVENT_DATE) as EVENT_DATE,SUM(AMOUNT) AMOUNT,SUM(AMOUNT_REAL) AMOUNT_REAL
  FROM (
	  SELECT BUDGET.EVENT_DATE,ISNULL(BUDGET.AMOUNT,0) AMOUNT,ISNULL(REALAMT.AMOUNT_REAL,0) AMOUNT_REAL FROM (
	  	SELECT --工地預算
	  	 CONVERT (Datetime, CONVERT(varchar, r.BUDGET_YEAR) + '/' + CONVERT(varchar, r.BUDGET_MONTH) 
		 + '/' + CONVERT(varchar, S.BUDGET_DAY)) EVENT_DATE,
	  	 SUBJECT_ID, BUDGET_YEAR AS CURRENT_YEAR, BUDGET_MONTH, AMOUNT 
	  	FROM PLAN_SITE_BUDGET R  JOIN FIN_SUBJECT S ON  R.SUBJECT_ID=S.FIN_SUBJECT_ID
	  	WHERE S.BUDGET_DAY IS NOT NULL 
	  	UNION
	  	SELECT 
	  	 Dateadd(day,-1,Dateadd(month,1,CONVERT (Datetime,CONVERT(varchar, r.BUDGET_YEAR) + '/' 
		 + CONVERT(varchar, r.BUDGET_MONTH) + '/' + CONVERT(varchar, 1))))  EVENT_DATE, 
	  	 SUBJECT_ID, BUDGET_YEAR AS CURRENT_YEAR, BUDGET_MONTH, AMOUNT 
	  	FROM PLAN_SITE_BUDGET R  JOIN FIN_SUBJECT S ON  R.SUBJECT_ID=S.FIN_SUBJECT_ID
	  	WHERE S.BUDGET_DAY IS NULL 
       ) BUDGET LEFT JOIN (
	   --工地費用
          SELECT PAYMENT_DATE EVENT_DATE,
          i.FIN_SUBJECT_ID SUBJECT_ID, sum(i.AMOUNT) as AMOUNT_REAL FROM FIN_EXPENSE_FORM f 
          LEFT JOIN FIN_EXPENSE_ITEM i ON f.EXP_FORM_ID = i.EXP_FORM_ID WHERE i.FIN_SUBJECT_ID Like '1%'
          AND f.STATUS = 30 GROUP BY PAYMENT_DATE,FIN_SUBJECT_ID
	   ) REALAMT ON BUDGET.EVENT_DATE=REALAMT.EVENT_DATE AND BUDGET.SUBJECT_ID=REALAMT.SUBJECT_ID 
	) V GROUP BY dbo.getWorkDay(v.EVENT_DATE)
),
OtherExpense AS (
 SELECT dbo.getWorkDay(v.EVENT_DATE) as EVENT_DATE,SUM(AMOUNT) AMOUNT,SUM(0) AMOUNT_REAL
 FROM (
	--其他帳務借款
  	SELECT ISNULL(EVENT_DATE,PAYBACK_DATE) EVENT_DATE,SUM(AMOUNT) AMOUNT,SUM(0) AMOUNT_REAL
  	FROM FIN_LOAN_TRANACTION WHERE 
  	BL_ID IN (SELECT BL_ID FROM FIN_BANK_LOAN WHERE IS_SUPPLIER='Y')
  	AND TRANSACTION_TYPE=-1
  	GROUP BY ISNULL(EVENT_DATE,PAYBACK_DATE) 
    UNION --廠商請款
  	SELECT PAYMENT_DATE EVENT_DATE,SUM(ISNULL(AMOUNT_PAID,AMOUNT_PAYABLE)) AMOUNT,SUM(0) AMOUNT_REAL FROM PLAN_ACCOUNT
  	WHERE ISDEBIT = 'N' AND ISNULL(STATUS, 10) <> 0 
  	GROUP BY PAYMENT_DATE
    UNION --銀行還款
 	SELECT ISNULL(EVENT_DATE,PAYBACK_DATE) EVENT_DATE,SUM(AMOUNT) AMOUNT,SUM(0) AMOUNT_REAL
  	FROM FIN_LOAN_TRANACTION WHERE 
  	BL_ID IN (SELECT BL_ID FROM FIN_BANK_LOAN WHERE ISNULL(IS_SUPPLIER,'N')='N') AND REMARK NOT LIKE ('備償款%')
	AND TRANSACTION_TYPE=1
  	GROUP BY ISNULL(EVENT_DATE,PAYBACK_DATE) 
 ) V GROUP BY dbo.getWorkDay(v.EVENT_DATE)
),
Outcome AS (
 SELECT EVENT_DATE ,SUM(AMOUNT) AMOUNT,SUM(0) AMOUNT_REAL FROM 
 (
  SELECT * FROM  OtherExpense
   UNION 
     SELECT * FROM SiteExpense
	    UNION 
     SELECT * FROM CompanyExpense
	 ) v GROUP BY EVENT_DATE
),
Income AS (
	--調整星期六,日至工作日 
     SELECT dbo.getWorkDay(v.EVENT_DATE) as EVENT_DATE,SUM(AMOUNT) AMOUNT
	 FROM (
	 --廠商還款
	 SELECT ISNULL(EVENT_DATE,PAYBACK_DATE) EVENT_DATE,SUM(AMOUNT) AMOUNT 
     FROM FIN_LOAN_TRANACTION WHERE 
     BL_ID IN (SELECT BL_ID FROM FIN_BANK_LOAN WHERE IS_SUPPLIER='Y')
     AND TRANSACTION_TYPE=1
     GROUP BY ISNULL(EVENT_DATE,PAYBACK_DATE) 
     UNION --業主請款
     SELECT PAYMENT_DATE,SUM(ISNULL(AMOUNT_PAID,AMOUNT_PAYABLE)) INAMOUNT FROM PLAN_ACCOUNT
      WHERE ISDEBIT = 'Y' AND ISNULL(STATUS, 10) <> 0 
     GROUP BY PAYMENT_DATE
	 UNION --銀行借款流入
	SELECT ISNULL(EVENT_DATE,PAYBACK_DATE) EVENT_DATE,SUM(AMOUNT) AMOUNT
  	FROM FIN_LOAN_TRANACTION WHERE 
  	BL_ID IN (SELECT BL_ID FROM FIN_BANK_LOAN WHERE ISNULL(IS_SUPPLIER,'N')='N') AND REMARK NOT LIKE ('備償款%')
	AND TRANSACTION_TYPE=-1
  	GROUP BY ISNULL(EVENT_DATE,PAYBACK_DATE) 
	 ) v GROUP BY dbo.getWorkDay(v.EVENT_DATE)
),
LoanQuota AS (
-- 可用貸款
SELECT SUM(AMOUNT) AMT
FROM(
 SELECT 
 IIF(vaRatio>ISNULL(CUM_AR_RATIO,0),QUOTA+SumTransactionAmount,((QUOTA * QUOTA_AVAILABLE_RATIO / 100 )
 + SumTransactionAmount)) AMOUNT
	 FROM (
	 SELECT B.QUOTA,B.DUE_DATE, P.PROJECT_NAME,B.QUOTA_AVAILABLE_RATIO, B.CUM_AR_RATIO,
	 ISNULL(ROUND(V.vaRatio * 100, 0),0) AS vaRatio  , 
     (SELECT ISNULL(IIF(QUOTA_RECYCLABLE = 'Y', SUM(TRANSACTION_TYPE*AMOUNT),(SUM(IIF(TRANSACTION_TYPE = 1,0,-1)*AMOUNT))), 0) 
      FROM FIN_LOAN_TRANACTION T WHERE T.BL_ID = B.BL_ID) SumTransactionAmount 
      FROM FIN_BANK_LOAN B LEFT JOIN TND_PROJECT P ON B.PROJECT_ID = P.PROJECT_ID LEFT JOIN
      (SELECT va.PROJECT_ID AS PROJECT_ID, ISNULL(va.VALUATION_AMOUNT, 0) / ISNULL(pi.contractAtm, 1) AS vaRatio 
      FROM (SELECT PROJECT_ID, SUM(VALUATION_AMOUNT) AS VALUATION_AMOUNT 
      FROM PLAN_VALUATION_FORM GROUP BY PROJECT_ID)va LEFT JOIN
      (SELECT PROJECT_ID, SUM(ITEM_UNIT_PRICE * ITEM_QUANTITY) AS contractAtm 
      FROM PLAN_ITEM GROUP BY PROJECT_ID)pi ON va.PROJECT_ID = pi.PROJECT_ID)V 
	  ON B.PROJECT_ID = V.PROJECT_ID WHERE ISNULL(IS_SUPPLIER, 'N') <> 'Y' AND B.DUE_DATE>GETDATE()
	) LoanDetail
 ) LoanSum
)
--輸出結果
SELECT DATE_CASHFLOW,SUM(AMOUNT_BANK) AMOUNT_BANK,SUM(AMOUNT_INFLOW) AMOUNT_INFLOW,
Sum(AMOUNT_OUTFLOW) AMOUNT_OUTFLOW,
(SELECT AMT FROM LoanQuota) LOAN_QUOTA_RUNNING_TOTAL
,SUM(SUM(AMOUNT_BANK)+ SUM(AMOUNT_INFLOW)-Sum(AMOUNT_OUTFLOW)) OVER(ORDER BY DATE_CASHFLOW
          ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW) 
               AS CASH_RUNNING_TOTAL
 FROM (
---銀行存款
  SELECT CONVERT(varchar(10), GETDATE() , 111) DATE_CASHFLOW,SUM(CUR_AMOUNT) AMOUNT_BANK,
    0 AMOUNT_INFLOW,0 AMOUNT_OUTFLOW,0.AMOUNT_REAL 
  FROM FIN_BANK_ACCOUNT
  UNION 
    --彙整支出/收入資料   
    SELECT  CONVERT(varchar(10), dbo.getWorkDay(v.EVENT_DATE) , 111) DATE_CASHFLOW,0 AMOUNT_BANK,
	  ISNULL(i.AMOUNT,0) AMOUNT_INFLOW, ISNULL(o.AMOUNT,0) AMOUNT_OUTFLOW,ISNULL(o.AMOUNT_REAL,0) AMOUNT_REAL
       from vw_TimeIndex4CashFlow v 
	   LEFT JOIN Outcome o ON  dbo.getWorkDay(v.EVENT_DATE)=dbo.getWorkDay(o.EVENT_DATE )
	   LEFT JOIN Income i ON dbo.getWorkDay(v.EVENT_DATE)=dbo.getWorkDay(i.EVENT_DATE)
	   ) B  GROUP BY DATE_CASHFLOW ORDER BY DATE_CASHFLOW
";
                logger.Debug("SQL=" + sql);
                lstItem = context.Database.SqlQuery<CashFlowFunction>(sql).ToList();
                logger.Info("Get Cash Flow Count=" + lstItem.Count);
            }

            return lstItem;
        }
        //取得特定日期之收入(Debit)
        public List<PLAN_ACCOUNT> getDebitByDate(string date)
        {
            List<PLAN_ACCOUNT> lstDebit = new List<PLAN_ACCOUNT>();
            using (var context = new topmepEntities())
            {
                lstDebit = context.PLAN_ACCOUNT.SqlQuery("select a.* from PLAN_ACCOUNT a "
                    + "where a.ISDEBIT = 'Y' AND CONVERT(char(10),a.PAYMENT_DATE,111) = @date "
                   , new SqlParameter("date", date)).ToList();
            }
            return lstDebit;
        }

        //取得特定日期之支出(Credit)
        public List<CashFlowBalance> getCreditByDate(string date)
        {
            List<CashFlowBalance> lstCredit = new List<CashFlowBalance>();
            int plan_account_id = 0;
            string account_type = "L";
            using (var context = new topmepEntities())
            {
                lstCredit = context.Database.SqlQuery<CashFlowBalance>("select a.PLAN_ACCOUNT_ID, a.ACCOUNT_TYPE, a.ACCOUNT_FORM_ID, a.PAYMENT_DATE, a.AMOUNT_PAYABLE, a.AMOUNT_PAID, a.STATUS from PLAN_ACCOUNT a " +
                    "where a.ISDEBIT = 'N' AND CONVERT(char(10), a.PAYMENT_DATE, 111) = @date UNION select " + plan_account_id + ", '" + account_type + "', CONVERT(VARCHAR,flt.TID),flt.PAYBACK_DATE, " +
                    "flt.AMOUNT, flt.TRANSACTION_TYPE  from FIN_LOAN_TRANACTION flt where CONVERT(char(10),flt.PAYBACK_DATE,111) =@date "
                   , new SqlParameter("date", date)).ToList();
            }
            return lstCredit;
        }

        #region 公司費用預算
        //公司費用預算上傳 
        public int refreshExpBudget(List<FIN_EXPENSE_BUDGET> items)
        {
            int i = 0;
            logger.Info("refreshExpBudgetItem = " + items.Count);
            //2.將Excel 資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (FIN_EXPENSE_BUDGET item in items)
                {
                    context.FIN_EXPENSE_BUDGET.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add FIN_EXPENSE_BUDGET count =" + i);
            return i;
        }
        public int delExpBudgetByYear(int year)
        {
            logger.Info("remove all expense budget by budget year=" + year);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete all FIN_EXPENSE_BUDGET by budget year =" + year);
                i = context.Database.ExecuteSqlCommand("DELETE FROM FIN_EXPENSE_BUDGET WHERE BUDGET_YEAR=@year", new SqlParameter("year", year));
            }
            logger.Debug("deleteFIN_EXPENSE_BUDGET count=" + i);
            return i;
        }

        //取得特定年度公司費用預算
        public List<ExpenseBudgetSummary> getExpBudgetByYear(int year)
        {
            List<ExpenseBudgetSummary> lstExpBudget = new List<ExpenseBudgetSummary>();
            using (var context = new topmepEntities())
            {
                string sql = @"--預算數
                            SELECT ROW_NUMBER() OVER(ORDER BY A.SUBJECT_ID) AS SUB_NO, A.*, 
                            SUM(ISNULL(A.JAN,0))+ SUM(ISNULL(A.FEB,0))+ SUM(ISNULL(A.MAR,0)) + SUM(ISNULL(A.APR,0)) + SUM(ISNULL(A.MAY,0)) + SUM(ISNULL(A.JUN,0)) 
                            + SUM(ISNULL(A.JUL, 0)) + SUM(ISNULL(A.AUG, 0)) + SUM(ISNULL(A.SEP, 0)) + SUM(ISNULL(A.OCT, 0)) + SUM(ISNULL(A.NOV, 0)) + SUM(ISNULL(A.DEC, 0)) AS HTOTAL 
                            FROM(SELECT SUBJECT_NAME, SUBJECT_ID, [01] As 'JAN', [02] As 'FEB', [03] As 'MAR', [04] As 'APR', [05] As 'MAY', [06] As 'JUN', [07] As 'JUL', [08] As 'AUG', [09] As 'SEP', [10] As 'OCT', [11] As 'NOV', [12] As 'DEC' 
                            FROM(
                            SELECT FIN_SUBJECT_ID SUBJECT_ID,BUDGET_MONTH ,AMOUNT AMOUNT, BUDGET_YEAR,SUBJECT_NAME
                            FROM 
                            (SELECT * FROM FIN_SUBJECT  WHERE CATEGORY='公司營業費用' ) SUBJ
                            LEFT OUTER JOIN 
                            (SELECT BUDGET_YEAR,BUDGET_MONTH,SUBJECT_ID,AMOUNT,CURRENT_YEAR FROM FIN_EXPENSE_BUDGET
                            WHERE BUDGET_YEAR=@year) BUD
                            ON SUBJ.FIN_SUBJECT_ID=BUD.SUBJECT_ID
                            ) As STable 
                            PIVOT(SUM(AMOUNT) FOR BUDGET_MONTH IN([01], [02], [03], [04], [05], [06], [07], [08], [09], [10], [11], [12])) As PTable)A 
                            GROUP BY A.SUBJECT_NAME, A.SUBJECT_ID, A.JAN, A.FEB, A.MAR,A.APR, A.MAY, A.JUN, A.JUL, A.AUG, A.SEP, A.OCT, A.NOV, A.DEC;";
                lstExpBudget = context.Database.SqlQuery<ExpenseBudgetSummary>(sql, new SqlParameter("year", year)).ToList();
            }
            return lstExpBudget;
        }

        //取得公司費用科目代碼與名稱
        public List<FIN_SUBJECT> getExpBudgetSubject()
        {
            List<FIN_SUBJECT> lstSubject = new List<FIN_SUBJECT>();
            using (var context = new topmepEntities())
            {
                lstSubject = context.Database.SqlQuery<FIN_SUBJECT>("SELECT * FROM FIN_SUBJECT WHERE CATEGORY = '公司營業費用' ORDER BY FIN_SUBJECT_ID ; ").ToList();
            }
            return lstSubject;
        }



        //取得特定年度之公司費用總執行金額
        public ExpenseBudgetSummary getTotalOperationExpAmount(int year)
        {
            ExpenseBudgetSummary lstExpAmount = null;
            using (var context = new topmepEntities())
            {
                lstExpAmount = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT SUM(ei.AMOUNT) AS TOTAL_OPERATION_EXP FROM FIN_EXPENSE_ITEM ei LEFT JOIN FIN_EXPENSE_FORM ef " +
                    "ON ei.EXP_FORM_ID = ef.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs ON ei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID WHERE fs.CATEGORY = '公司營業費用' " +
                    "AND ef.OCCURRED_YEAR = @year AND ef.OCCURRED_MONTH > 6 OR fs.CATEGORY = '公司營業費用' AND ef.OCCURRED_YEAR = @year + 1 AND ef.OCCURRED_MONTH < 7  "
               , new SqlParameter("year", year)).FirstOrDefault();
            }
            return lstExpAmount;
        }

        //取得特定年度公司費用彙整
        public List<ExpenseBudgetSummary> getExpSummaryByYear(int year)
        {
            List<ExpenseBudgetSummary> lstExpBudget = new List<ExpenseBudgetSummary>();
            using (var context = new topmepEntities())
            {
                string sql = @"SELECT ROW_NUMBER() OVER(ORDER BY C.FIN_SUBJECT_ID) + 1 AS SUB_NO, C.*, SUM(ISNULL(C.JAN, 0)) + SUM(ISNULL(C.FEB, 0)) + SUM(ISNULL(C.MAR, 0)) + SUM(ISNULL(C.APR, 0)) + SUM(ISNULL(C.MAY, 0)) + SUM(ISNULL(C.JUN, 0)) 
                        + SUM(ISNULL(C.JUL, 0)) + SUM(ISNULL(C.AUG, 0)) + SUM(ISNULL(C.SEP, 0)) + SUM(ISNULL(C.OCT, 0)) + SUM(ISNULL(C.NOV, 0)) + SUM(ISNULL(C.DEC, 0)) AS HTOTAL 
                        FROM (
                        SELECT SUBJECT_NAME, FIN_SUBJECT_ID, [01] As 'JAN', [02] As 'FEB', [03] As 'MAR', [04] As 'APR', [05] As 'MAY', [06] As 'JUN', [07] As 'JUL', [08] As 'AUG', [09] As 'SEP', [10] As 'OCT', [11] As 'NOV', [12] As 'DEC' 
                        FROM (
                        SELECT B.OCCURRED_MONTH, fs.FIN_SUBJECT_ID, fs.SUBJECT_NAME, B.AMOUNT FROM FIN_SUBJECT fs LEFT JOIN(SELECT ef.OCCURRED_MONTH, ei.FIN_SUBJECT_ID, ei.AMOUNT FROM FIN_EXPENSE_ITEM ei 
                        LEFT JOIN FIN_EXPENSE_FORM ef ON ei.EXP_FORM_ID = ef.EXP_FORM_ID WHERE ef.OCCURRED_YEAR = @year AND ef.OCCURRED_MONTH > 6 OR ef.OCCURRED_YEAR = @year + 1 AND ef.OCCURRED_MONTH < 7)B 
                        ON fs.FIN_SUBJECT_ID = B.FIN_SUBJECT_ID WHERE fs.CATEGORY = '公司營業費用') As STable 
                        PIVOT(SUM(AMOUNT) FOR OCCURRED_MONTH IN([01], [02], [03], [04], [05], [06], [07], [08], [09], [10], [11], [12])) As PTable) C 
                        GROUP BY C.SUBJECT_NAME, C.FIN_SUBJECT_ID, C.JAN, C.FEB, C.MAR, C.APR, C.MAY, C.JUN, C.JUL, C.AUG, C.SEP, C.OCT, C.NOV, C.DEC";
                lstExpBudget = context.Database.SqlQuery<ExpenseBudgetSummary>(sql, new SqlParameter("year", year)).ToList();
            }
            return lstExpBudget;
        }

        //取得特定年度公司費用每月預算總和
        public List<ExpenseBudgetByMonth> getExpBudgetOfMonthByYear(int year)
        {
            List<ExpenseBudgetByMonth> lstExpBudget = new List<ExpenseBudgetByMonth>();
            using (var context = new topmepEntities())
            {
                lstExpBudget = context.Database.SqlQuery<ExpenseBudgetByMonth>("SELECT SUM(F.JAN) AS JAN, SUM(F.FEB) AS FEB, SUM(F.MAR) AS MAR, SUM(F.APR) AS APR, SUM(F.MAY) AS MAY, SUM(F.JUN) AS JUN, " +
                   "SUM(F.JUL) AS JUL, SUM(F.AUG) AS AUG, SUM(F.SEP) AS SEP, SUM(F.OCT) AS OCT, SUM(F.NOV) AS NOV, SUM(F.DEC) AS DEC, SUM(F.HTOTAL) AS HTOTAL " +
                   "FROM (SELECT A.*, SUM(ISNULL(A.JAN,0))+ SUM(ISNULL(A.FEB,0))+ SUM(ISNULL(A.MAR,0)) + SUM(ISNULL(A.APR,0)) + SUM(ISNULL(A.MAY,0)) + SUM(ISNULL(A.JUN,0)) " +
                   "+ SUM(ISNULL(A.JUL, 0)) + SUM(ISNULL(A.AUG, 0)) + SUM(ISNULL(A.SEP, 0)) + SUM(ISNULL(A.OCT, 0)) + SUM(ISNULL(A.NOV, 0)) + SUM(ISNULL(A.DEC, 0)) AS HTOTAL " +
                   "FROM(SELECT SUBJECT_NAME, SUBJECT_ID, [01] As 'JAN', [02] As 'FEB', [03] As 'MAR', [04] As 'APR', [05] As 'MAY', [06] As 'JUN', [07] As 'JUL', [08] As 'AUG', [09] As 'SEP', [10] As 'OCT', [11] As 'NOV', [12] As 'DEC' " +
                   "FROM(SELECT eb.SUBJECT_ID, eb.BUDGET_MONTH, eb.AMOUNT, eb.BUDGET_YEAR, fs.SUBJECT_NAME FROM FIN_EXPENSE_BUDGET eb LEFT JOIN FIN_SUBJECT fs ON eb.SUBJECT_ID = fs.FIN_SUBJECT_ID WHERE BUDGET_YEAR = @year) As STable " +
                   "PIVOT(SUM(AMOUNT) FOR BUDGET_MONTH IN([01], [02], [03], [04], [05], [06], [07], [08], [09], [10], [11], [12])) As PTable)A " +
                   "GROUP BY A.SUBJECT_NAME, A.SUBJECT_ID, A.JAN, A.FEB, A.MAR,A.APR, A.MAY, A.JUN, A.JUL, A.AUG, A.SEP, A.OCT, A.NOV, A.DEC)F ; "
                   , new SqlParameter("year", year)).ToList();
            }
            return lstExpBudget;
        }

        #endregion

        #region 公司費用
        public ExpenseFormFunction formEXP = null;
        public List<ExpenseBudgetSummary> EXPItem = null;
        public List<ExpenseBudgetSummary> siteEXPItem = null;
        public ExpenseBudgetSummary ExpAmt = null;
        public ExpenseBudgetSummary EarlyCumAmt = null;
        public ExpenseBudgetSummary SiteEarlyCumAmt = null;
        //取得公司費用項目
        public List<FIN_SUBJECT> getSubjectOfExpense()
        {
            List<FIN_SUBJECT> lstSubject = new List<FIN_SUBJECT>();
            using (var context = new topmepEntities())
            {
                lstSubject = context.Database.SqlQuery<FIN_SUBJECT>("SELECT * FROM FIN_SUBJECT WHERE CATEGORY = '公司營業費用' ORDER BY FIN_SUBJECT_ID ; ").ToList();
                logger.Info("Get Subject of Operating Expense Count=" + lstSubject.Count);
            }
            return lstSubject;
        }

        //取得特定費用項目
        public List<FIN_SUBJECT> getSubjectByChkItem(string[] lstItemId)
        {
            List<FIN_SUBJECT> lstSubject = new List<FIN_SUBJECT>();
            string ItemId = "";
            for (i = 0; i < lstItemId.Count(); i++)
            {
                if (i < lstItemId.Count() - 1)
                {
                    ItemId = ItemId + "'" + lstItemId[i] + "'" + ",";
                }
                else
                {
                    ItemId = ItemId + "'" + lstItemId[i] + "'";
                }
            }
            using (var context = new topmepEntities())
            {
                lstSubject = context.Database.SqlQuery<FIN_SUBJECT>("SELECT * FROM FIN_SUBJECT WHERE FIN_SUBJECT_ID IN (" + ItemId + ") ; ").ToList();
                logger.Info("Get Subject of Expense  Count=" + lstSubject.Count);
            }
            return lstSubject;
        }

        public string newExpenseForm(FIN_EXPENSE_FORM form)
        {
            //1.建立公司營業費用單/工地費用單
            logger.Info("create new expense form ");
            string sno_key = "EXP";
            SerialKeyService snoservice = new SerialKeyService();
            form.EXP_FORM_ID = snoservice.getSerialKey(sno_key);
            logger.Info("new expense form =" + form.ToString());
            using (var context = new topmepEntities())
            {
                context.FIN_EXPENSE_FORM.Add(form);
                int i = context.SaveChanges();
                logger.Debug("Add form=" + i);
                logger.Info("expense form id = " + form.EXP_FORM_ID);
                //if (i > 0) { status = true; };
            }
            return form.EXP_FORM_ID;
        }

        public int AddExpenseItems(List<FIN_EXPENSE_ITEM> lstItem)
        {
            //2.新增費用項目資料
            int j = 0;
            logger.Info("add expense items = " + lstItem.Count);
            using (var context = new topmepEntities())
            {
                //3.將費用項目資料寫入 
                foreach (FIN_EXPENSE_ITEM item in lstItem)
                {
                    context.FIN_EXPENSE_ITEM.Add(item);
                }

                j = context.SaveChanges();
                logger.Info("add expense count =" + j);
            }
            return j;
        }

        //取得公司費用單/工地費用單
        public void getEXPByExpId(string expid)
        {
            logger.Info("get form : formid=" + expid);
            using (var context = new topmepEntities())
            {
                //取得費用單檔頭資訊
                string sql = "SELECT fef.EXP_FORM_ID, fef.PROJECT_ID, fef.OCCURRED_YEAR, fef.OCCURRED_MONTH, fef.PAYMENT_DATE, fef.STATUS, fef.CREATE_ID, fef.CREATE_DATE, fef.REMARK, fef.MODIFY_DATE, fef.PASS_CREATE_ID, fef.PASS_CREATE_DATE, fef.APPROVE_CREATE_ID, fef.APPROVE_CREATE_DATE, " +
                    "fef.JOURNAL_CREATE_ID, fef.JOURNAL_CREATE_DATE, fef.PAYEE, p.PROJECT_NAME FROM FIN_EXPENSE_FORM fef LEFT JOIN TND_PROJECT p ON fef.PROJECT_ID = p.PROJECT_ID WHERE fef.EXP_FORM_ID =@expid ";
                formEXP = context.Database.SqlQuery<ExpenseFormFunction>(sql, new SqlParameter("expid", expid)).First();
                //取得公司營業費用單明細
                EXPItem = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT G.*, C.CUM_BUDGET, ISNULL(D.CUM_AMOUNT, 0) AS CUM_AMOUNT, ISNULL(G.AMOUNT, 0) + ISNULL(D.CUM_AMOUNT, 0) AS CUR_CUM_AMOUNT, (ISNULL(G.AMOUNT, 0) + ISNULL(D.CUM_AMOUNT, 0)) / IIF(G.BUDGET_AMOUNT IS NULL, 100, G.BUDGET_AMOUNT) * 100 AS CUR_CUM_RATIO, " +
                    "ROW_NUMBER() OVER(ORDER BY G.EXP_ITEM_ID) AS NO FROM (SELECT A.*, B.AMOUNT AS BUDGET_AMOUNT FROM (SELECT fei.*, fef.OCCURRED_YEAR, fef.OCCURRED_MONTH, fef.STATUS, fs.SUBJECT_NAME FROM FIN_EXPENSE_ITEM fei LEFT JOIN FIN_EXPENSE_FORM fef ON fei.EXP_FORM_ID = fef.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs " +
                    "ON fei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID WHERE fei.EXP_FORM_ID = @expid)A " +
                    "LEFT JOIN (SELECT F.*, feb.AMOUNT FROM (SELECT fei.FIN_SUBJECT_ID, fef.OCCURRED_YEAR, fef.OCCURRED_MONTH FROM FIN_EXPENSE_ITEM fei LEFT JOIN FIN_EXPENSE_FORM fef ON fei.EXP_FORM_ID = fef.EXP_FORM_ID WHERE fei.EXP_FORM_ID = @expid " +
                    "GROUP BY fei.FIN_SUBJECT_ID, fef.OCCURRED_YEAR, fef.OCCURRED_MONTH)F LEFT JOIN (SELECT SUBJECT_ID, SUM(AMOUNT) AS AMOUNT FROM FIN_EXPENSE_BUDGET WHERE BUDGET_YEAR = IIF(" + formEXP.OCCURRED_MONTH + "< 7," + (formEXP.OCCURRED_YEAR - 1) + "," + formEXP.OCCURRED_YEAR + ")" +
                    "GROUP BY SUBJECT_ID)feb ON F.FIN_SUBJECT_ID = feb.SUBJECT_ID)B ON A.FIN_SUBJECT_ID = B.FIN_SUBJECT_ID)G " +
                    "LEFT JOIN (SELECT SUBJECT_ID, SUM(AMOUNT) AS CUM_BUDGET FROM FIN_EXPENSE_BUDGET WHERE CURRENT_YEAR = " + formEXP.OCCURRED_YEAR + "AND BUDGET_MONTH <= IIF(" + formEXP.OCCURRED_MONTH + "< 7," + formEXP.OCCURRED_MONTH + ", 0) OR BUDGET_YEAR = " +
                    "IIF(" + formEXP.OCCURRED_MONTH + "< 7," + (formEXP.OCCURRED_YEAR - 1) + "," + formEXP.OCCURRED_YEAR + ") AND BUDGET_MONTH BETWEEN 7 AND IIF(" + formEXP.OCCURRED_MONTH + "< 7, 12," + formEXP.OCCURRED_MONTH + ") GROUP BY SUBJECT_ID) C ON G.FIN_SUBJECT_ID = C.SUBJECT_ID " +
                    "LEFT JOIN (SELECT fei.FIN_SUBJECT_ID, SUM(AMOUNT) AS CUM_AMOUNT FROM FIN_EXPENSE_ITEM fei LEFT JOIN FIN_EXPENSE_FORM fef ON fei.EXP_FORM_ID = fef.EXP_FORM_ID " +
                    "WHERE fef.OCCURRED_YEAR = " + formEXP.OCCURRED_YEAR + "AND fef.OCCURRED_MONTH <= IIF(" + formEXP.OCCURRED_MONTH + "< 7," + formEXP.OCCURRED_MONTH + ", 0) AND fef.STATUS >= 30 AND fef.CREATE_DATE <= CONVERT(VARCHAR,CONVERT(datetime, REPLACE(REPLACE('" + formEXP.CREATE_DATE + "','上午',''),'下午','')+case when charindex('上午','" + formEXP.CREATE_DATE + "')>0 then 'AM' when charindex('下午','" + formEXP.CREATE_DATE + "')>0 then 'PM' end),120) " +
                    "OR fef.OCCURRED_YEAR = IIF(" + formEXP.OCCURRED_MONTH + "< 7," + (formEXP.OCCURRED_YEAR - 1) + "," + (formEXP.OCCURRED_YEAR) + ") " +
                    "AND fef.OCCURRED_MONTH BETWEEN 7 AND IIF(" + formEXP.OCCURRED_MONTH + "< 7, 12," + formEXP.OCCURRED_MONTH + ") AND fef.STATUS >= 30 AND fef.CREATE_DATE <= CONVERT(VARCHAR,CONVERT(datetime, REPLACE(REPLACE('" + formEXP.CREATE_DATE + "','上午',''),'下午','')+case when charindex('上午','" + formEXP.CREATE_DATE + "')>0 then 'AM' when charindex('下午','" + formEXP.CREATE_DATE + "')>0 then 'PM' end),120) GROUP BY fei.FIN_SUBJECT_ID)D ON G.FIN_SUBJECT_ID = D.FIN_SUBJECT_ID ", new SqlParameter("expid", expid)).ToList();
                logger.Debug("get query year of operating expense:" + formEXP.OCCURRED_YEAR);
                logger.Debug("get operating expense item count:" + EXPItem.Count);
                ExpAmt = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT SUM(AMOUNT) AS AMOUNT FROM FIN_EXPENSE_ITEM fei LEFT JOIN FIN_EXPENSE_FORM ef ON fei.EXP_FORM_ID = ef.EXP_FORM_ID WHERE ef.EXP_FORM_ID = @expid ", new SqlParameter("expid", expid)).FirstOrDefault();
                EarlyCumAmt = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT SUM(AMOUNT) AS AMOUNT FROM FIN_SUBJECT fs LEFT JOIN FIN_EXPENSE_ITEM fei ON fs.FIN_SUBJECT_ID = fei.FIN_SUBJECT_ID LEFT JOIN FIN_EXPENSE_FORM fef ON fei.EXP_FORM_ID = fef.EXP_FORM_ID " +
                    "WHERE fs.CATEGORY = '公司營業費用' AND fef.OCCURRED_YEAR = " + formEXP.OCCURRED_YEAR + "AND fef.OCCURRED_MONTH <= IIF(" + formEXP.OCCURRED_MONTH + " < 7, " + formEXP.OCCURRED_MONTH + ", 0) AND fef.STATUS >= 30 AND fef.CREATE_DATE < CONVERT(VARCHAR,CONVERT(datetime, REPLACE(REPLACE('" + formEXP.CREATE_DATE + "','上午',''),'下午','')+case when charindex('上午','" + formEXP.CREATE_DATE + "')>0 then 'AM' when charindex('下午','" + formEXP.CREATE_DATE + "')>0 then 'PM' end),120) " +
                    "OR fs.CATEGORY = '公司營業費用' AND fef.OCCURRED_YEAR = IIF(" + formEXP.OCCURRED_MONTH + " < 7, " + (formEXP.OCCURRED_YEAR - 1) + ", " + (formEXP.OCCURRED_YEAR) + ") " +
                    "AND fef.OCCURRED_MONTH BETWEEN 7 AND IIF(" + formEXP.OCCURRED_MONTH + " < 7, 12, " + formEXP.OCCURRED_MONTH + ") AND fef.STATUS >= 30 AND fef.CREATE_DATE < CONVERT(VARCHAR,CONVERT(datetime, REPLACE(REPLACE('" + formEXP.CREATE_DATE + "','上午',''),'下午','')+case when charindex('上午','" + formEXP.CREATE_DATE + "')>0 then 'AM' when charindex('下午','" + formEXP.CREATE_DATE + "')>0 then 'PM' end),120) ").FirstOrDefault();
                logger.Debug("get query create date of operating expense:" + formEXP.CREATE_DATE);
                //取得工地費用單明細
                if (null != formEXP.PROJECT_ID && formEXP.PROJECT_ID != "")
                {
                    sql = "SELECT G.*, C.CUM_BUDGET, ISNULL(D.CUM_AMOUNT, 0) AS CUM_AMOUNT, ISNULL(G.AMOUNT, 0) + ISNULL(D.CUM_AMOUNT, 0) AS CUR_CUM_AMOUNT, (ISNULL(G.AMOUNT, 0) + ISNULL(D.CUM_AMOUNT, 0)) / IIF(G.BUDGET_AMOUNT IS NULL, 1, G.BUDGET_AMOUNT) *100 AS CUR_CUM_RATIO, " +
                       "ROW_NUMBER() OVER(ORDER BY G.EXP_ITEM_ID) AS NO FROM (SELECT A.*, B.BUDGET AS BUDGET_AMOUNT FROM (SELECT fei.*, fef.OCCURRED_YEAR, fef.OCCURRED_MONTH, fef.STATUS, fs.SUBJECT_NAME FROM FIN_EXPENSE_ITEM fei LEFT JOIN FIN_EXPENSE_FORM fef ON fei.EXP_FORM_ID = fef.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs " +
                       "ON fei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID WHERE fei.EXP_FORM_ID = @expid)A " +
                       "LEFT JOIN (SELECT F.*, psb.BUDGET FROM (SELECT fei.FIN_SUBJECT_ID, fef.OCCURRED_YEAR, fef.OCCURRED_MONTH, fef.PROJECT_ID FROM FIN_EXPENSE_ITEM fei LEFT JOIN FIN_EXPENSE_FORM fef ON fei.EXP_FORM_ID = fef.EXP_FORM_ID WHERE fei.EXP_FORM_ID = @expid " +
                       "GROUP BY fei.FIN_SUBJECT_ID, fef.OCCURRED_YEAR, fef.OCCURRED_MONTH, fef.PROJECT_ID)F LEFT JOIN (SELECT SUBJECT_ID, SUM(AMOUNT) BUDGET FROM PLAN_SITE_BUDGET WHERE PROJECT_ID = '" + formEXP.PROJECT_ID + "' GROUP BY SUBJECT_ID)psb ON F.FIN_SUBJECT_ID = psb.SUBJECT_ID)B ON A.FIN_SUBJECT_ID = B.FIN_SUBJECT_ID)G " +
                       "LEFT OUTER JOIN (SELECT SUBJECT_ID, SUM(AMOUNT) AS CUM_BUDGET FROM PLAN_SITE_BUDGET WHERE PROJECT_ID = '" + formEXP.PROJECT_ID + "' AND BUDGET_YEAR = " + formEXP.OCCURRED_YEAR + " AND BUDGET_MONTH <= " + formEXP.OCCURRED_MONTH + " OR PROJECT_ID = '" + formEXP.PROJECT_ID + "' AND BUDGET_YEAR < " + formEXP.OCCURRED_YEAR + " GROUP BY SUBJECT_ID) C ON G.FIN_SUBJECT_ID = C.SUBJECT_ID " +
                       "LEFT OUTER JOIN (SELECT fei.FIN_SUBJECT_ID, SUM(AMOUNT) AS CUM_AMOUNT FROM FIN_EXPENSE_ITEM fei LEFT JOIN FIN_EXPENSE_FORM fef ON fei.EXP_FORM_ID = fef.EXP_FORM_ID " +
                       "WHERE fef.PROJECT_ID = '" + formEXP.PROJECT_ID + "' AND fef.OCCURRED_YEAR = " + formEXP.OCCURRED_YEAR + " AND fef.OCCURRED_MONTH <=" + formEXP.OCCURRED_MONTH + " AND fef.STATUS >= 30 AND fef.CREATE_DATE <= CONVERT(VARCHAR,CONVERT(datetime, REPLACE(REPLACE('" + formEXP.CREATE_DATE + "','上午',''),'下午','')+case when charindex('上午','" + formEXP.CREATE_DATE + "')>0 then 'AM' when charindex('下午','" + formEXP.CREATE_DATE + "')>0 then 'PM' end),120) " +
                       "OR fef.PROJECT_ID = '" + formEXP.PROJECT_ID + "' AND fef.OCCURRED_YEAR < " + formEXP.OCCURRED_YEAR + " AND  " +
                       "fef.STATUS >= 30 AND fef.CREATE_DATE <= CONVERT(VARCHAR,CONVERT(datetime, REPLACE(REPLACE('" + formEXP.CREATE_DATE + "','上午',''),'下午','')+case when charindex('上午','" + formEXP.CREATE_DATE + "')>0 then 'AM' when charindex('下午','" + formEXP.CREATE_DATE + "')>0 then 'PM' end),120) GROUP BY fei.FIN_SUBJECT_ID)D ON G.FIN_SUBJECT_ID = D.FIN_SUBJECT_ID ";
                    siteEXPItem = context.Database.SqlQuery<ExpenseBudgetSummary>(sql, new SqlParameter("expid", expid)).ToList();
                    logger.Debug("get query year of plan site expense:" + formEXP.OCCURRED_YEAR);
                    logger.Debug("get plan site expense item count:" + siteEXPItem.Count);
                    SiteEarlyCumAmt = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT SUM(AMOUNT) AS AMOUNT FROM FIN_SUBJECT fs LEFT JOIN FIN_EXPENSE_ITEM fei ON fs.FIN_SUBJECT_ID = fei.FIN_SUBJECT_ID LEFT JOIN FIN_EXPENSE_FORM fef ON fei.EXP_FORM_ID = fef.EXP_FORM_ID " +
                        "WHERE fs.CATEGORY = '工地費用' AND fef.PROJECT_ID = '" + formEXP.PROJECT_ID + "' AND fef.OCCURRED_YEAR = " + formEXP.OCCURRED_YEAR + " AND fef.OCCURRED_MONTH <= " + formEXP.OCCURRED_MONTH + " AND fef.STATUS >= 30 AND fef.CREATE_DATE < CONVERT(VARCHAR,CONVERT(datetime, REPLACE(REPLACE('" + formEXP.CREATE_DATE + "','上午',''),'下午','')+case when charindex('上午','" + formEXP.CREATE_DATE + "')>0 then 'AM' when charindex('下午','" + formEXP.CREATE_DATE + "')>0 then 'PM' end),120) " +
                        "OR fs.CATEGORY = '工地費用' AND fef.PROJECT_ID = '" + formEXP.PROJECT_ID + "' AND fef.OCCURRED_YEAR < " + formEXP.OCCURRED_YEAR + " AND  " +
                        "fef.STATUS >= 30 AND fef.CREATE_DATE < CONVERT(VARCHAR,CONVERT(datetime, REPLACE(REPLACE('" + formEXP.CREATE_DATE + "','上午',''),'下午','')+case when charindex('上午','" + formEXP.CREATE_DATE + "')>0 then 'AM' when charindex('下午','" + formEXP.CREATE_DATE + "')>0 then 'PM' end),120) ").FirstOrDefault();
                }
            }
        }

        public int refreshEXPForm(string formid, FIN_EXPENSE_FORM ef, List<FIN_EXPENSE_ITEM> lstItem)
        {
            logger.Info("Update expense form id =" + formid);
            int i = 0;
            int j = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.Entry(ef).State = EntityState.Modified;
                    i = context.SaveChanges();
                    logger.Debug("Update expense form =" + i);
                    logger.Info("expense form item = " + lstItem.Count);
                    //2.將item資料寫入 
                    foreach (FIN_EXPENSE_ITEM item in lstItem)
                    {
                        FIN_EXPENSE_ITEM existItem = null;
                        logger.Debug("form item id=" + item.EXP_ITEM_ID);
                        if (item.EXP_ITEM_ID != 0)
                        {
                            existItem = context.FIN_EXPENSE_ITEM.Find(item.EXP_ITEM_ID);
                        }
                        else
                        {
                            var parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("formid", formid));
                            parameters.Add(new SqlParameter("itemid", item.FIN_SUBJECT_ID));
                            string sql = "SELECT * FROM FIN_EXPENSE_ITEM WHERE EXP_FORM_ID=@formid AND FIN_SUBJECT_ID=@itemid";
                            logger.Info(sql + " ;" + formid + ",fin_subject_id=" + item.FIN_SUBJECT_ID);
                            FIN_EXPENSE_ITEM excelItem = context.FIN_EXPENSE_ITEM.SqlQuery(sql, parameters.ToArray()).First();
                            existItem = context.FIN_EXPENSE_ITEM.Find(excelItem.EXP_ITEM_ID);

                        }
                        logger.Debug("find exist item=" + existItem.FIN_SUBJECT_ID);
                        existItem.AMOUNT = item.AMOUNT;
                        existItem.ITEM_REMARK = item.ITEM_REMARK;
                        existItem.ITEM_QUANTITY = item.ITEM_QUANTITY;
                        existItem.ITEM_UNIT = item.ITEM_UNIT;
                        existItem.ITEM_UNIT_PRICE = item.ITEM_UNIT_PRICE;
                        context.FIN_EXPENSE_ITEM.AddOrUpdate(existItem);
                    }
                    j = context.SaveChanges();
                    logger.Debug("Update expense form item =" + j);
                    return j;
                }
                catch (Exception e)
                {
                    logger.Error("update new expense form id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }

        //更新費用單狀態為送審
        public int RefreshEXPStatusById(string expid)
        {
            int i = 0;
            logger.Info("update the status of EXP form by expid" + expid);
            string sql = "UPDATE  FIN_EXPENSE_FORM SET STATUS = 20 WHERE EXP_FORM_ID = @expid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("expid", expid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }

        //取得符合條件之公司營業費用單名單
        public List<OperatingExpenseFunction> getEXPListByExpId(string occurreddate, string subjectname, string expid, int status, string projectid)
        {
            logger.Info("search expense form by " + occurreddate + ", 費用單編號 =" + expid + ", 項目名稱 =" + subjectname + ", 估驗單狀態 =" + status + ", 專案編號 =" + projectid);
            List<OperatingExpenseFunction> lstForm = new List<OperatingExpenseFunction>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            if (15 == status)//預設狀態(STATUS >10 AND <20, 即狀態為退件或草稿)
            {
                string sql = "SELECT B.EXP_FORM_ID, B.PAYMENT_DATE, B.STATUS, left(B.Subjects,len(B.Subjects)-1) AS SUBJECT_NAME, " +
                    "CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) AS OCCURRED_DATE, ROW_NUMBER() OVER(ORDER BY B.EXP_FORM_ID) AS NO " +
                    "FROM(SELECT ef.*, (SELECT DISTINCT(cast( SUBJECT_NAME AS NVARCHAR ) + ',') from (SELECT ef.EXP_FORM_ID, fs.SUBJECT_NAME FROM FIN_EXPENSE_FORM ef " +
                    "LEFT JOIN FIN_EXPENSE_ITEM ei ON ef.EXP_FORM_ID = ei.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs ON ei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID)A " +
                    "WHERE ef.EXP_FORM_ID = A.EXP_FORM_ID FOR XML PATH('')) as Subjects FROM FIN_EXPENSE_FORM ef)B ";

                sql = sql + "WHERE B.STATUS > 10 AND ISNULL(B.PROJECT_ID, '') =@id ";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("id", projectid));
                // 費用發生年月條件
                if (null != occurreddate && occurreddate != "")
                {
                    sql = sql + "AND CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) =@occurreddate ";
                    parameters.Add(new SqlParameter("occurreddate", occurreddate));
                }
                //費用單編號條件
                if (null != expid && expid != "")
                {
                    sql = sql + "AND B.EXP_FORM_ID =@expid ";
                    parameters.Add(new SqlParameter("expid", expid));
                }
                //項目名稱條件
                if (null != subjectname && subjectname != "")
                {
                    sql = sql + "AND Subjects LIKE @subjectname ";
                    parameters.Add(new SqlParameter("subjectname", '%' + subjectname + '%'));
                }
                using (var context = new topmepEntities())
                {
                    logger.Debug("get expense form sql=" + sql);
                    lstForm = context.Database.SqlQuery<OperatingExpenseFunction>(sql, parameters.ToArray()).ToList();
                }
                logger.Info("get expense form count=" + lstForm.Count);
            }
            else if (20 == status)
            {
                string sql = "SELECT B.EXP_FORM_ID, B.PAYMENT_DATE, B.STATUS, left(B.Subjects,len(B.Subjects)-1) AS SUBJECT_NAME, " +
                    "CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) AS OCCURRED_DATE, ROW_NUMBER() OVER(ORDER BY B.EXP_FORM_ID) AS NO " +
                    "FROM(SELECT ef.*,(SELECT cast( SUBJECT_NAME AS NVARCHAR ) + ',' from (SELECT ef.EXP_FORM_ID, fs.SUBJECT_NAME FROM FIN_EXPENSE_FORM ef " +
                    "LEFT JOIN FIN_EXPENSE_ITEM ei ON ef.EXP_FORM_ID = ei.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs ON ei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID)A " +
                    "WHERE ef.EXP_FORM_ID = A.EXP_FORM_ID FOR XML PATH('')) as Subjects FROM FIN_EXPENSE_FORM ef)B ";

                sql = sql + "WHERE B.STATUS = 20 AND ISNULL(B.PROJECT_ID, '') =@id ";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("id", projectid));
                // 費用發生年月條件
                if (null != occurreddate && occurreddate != "")
                {
                    sql = sql + "AND CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) =@occurreddate ";
                    parameters.Add(new SqlParameter("occurreddate", occurreddate));
                }
                //費用單編號條件
                if (null != expid && expid != "")
                {
                    sql = sql + "AND B.EXP_FORM_ID =@expid ";
                    parameters.Add(new SqlParameter("expid", expid));
                }
                //項目名稱條件
                if (null != subjectname && subjectname != "")
                {
                    sql = sql + "AND Subjects LIKE @subjectname ";
                    parameters.Add(new SqlParameter("subjectname", '%' + subjectname + '%'));
                }
                using (var context = new topmepEntities())
                {
                    logger.Debug("get expense form sql=" + sql);
                    lstForm = context.Database.SqlQuery<OperatingExpenseFunction>(sql, parameters.ToArray()).ToList();
                }
                logger.Info("get expense form count=" + lstForm.Count);
            }
            else if (30 == status)
            {
                string sql = "SELECT ISNULL(B.PROJECT_ID,'公司營業費用') AS PROJECT_ID, B.PROJECT_NAME, B.EXP_FORM_ID, B.PAYMENT_DATE, B.STATUS, left(B.Subjects,len(B.Subjects)-1) AS SUBJECT_NAME, " +
                    "CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) AS OCCURRED_DATE, ROW_NUMBER() OVER(ORDER BY B.EXP_FORM_ID) AS NO " +
                    "FROM(SELECT ef.*, p.PROJECT_NAME, (SELECT cast( SUBJECT_NAME AS NVARCHAR ) + ',' from (SELECT ef.EXP_FORM_ID, fs.SUBJECT_NAME FROM FIN_EXPENSE_FORM ef " +
                    "LEFT JOIN FIN_EXPENSE_ITEM ei ON ef.EXP_FORM_ID = ei.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs ON ei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID)A " +
                    "WHERE ef.EXP_FORM_ID = A.EXP_FORM_ID FOR XML PATH('')) as Subjects FROM FIN_EXPENSE_FORM ef LEFT JOIN TND_PROJECT p ON ef.PROJECT_ID = p.PROJECT_ID)B ";

                sql = sql + "WHERE B.STATUS = 30 ";
                var parameters = new List<SqlParameter>();
                // 費用發生年月條件
                if (null != occurreddate && occurreddate != "")
                {
                    sql = sql + "AND CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) =@occurreddate ";
                    parameters.Add(new SqlParameter("occurreddate", occurreddate));
                }
                //費用單編號條件
                if (null != expid && expid != "")
                {
                    sql = sql + "AND B.EXP_FORM_ID =@expid ";
                    parameters.Add(new SqlParameter("expid", expid));
                }
                //項目名稱條件
                if (null != subjectname && subjectname != "")
                {
                    sql = sql + "AND Subjects LIKE @subjectname ";
                    parameters.Add(new SqlParameter("subjectname", '%' + subjectname + '%'));
                }
                using (var context = new topmepEntities())
                {
                    logger.Debug("get expense form sql=" + sql);
                    lstForm = context.Database.SqlQuery<OperatingExpenseFunction>(sql, parameters.ToArray()).ToList();
                }
                logger.Info("get expense form count=" + lstForm.Count);
            }
            else if (40 == status)
            {
                string sql = "SELECT B.EXP_FORM_ID, B.PAYMENT_DATE, B.STATUS, left(B.Subjects,len(B.Subjects)-1) AS SUBJECT_NAME, " +
                    "CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) AS OCCURRED_DATE, ROW_NUMBER() OVER(ORDER BY B.EXP_FORM_ID) AS NO " +
                    "FROM(SELECT ef.*,(SELECT cast( SUBJECT_NAME AS NVARCHAR ) + ',' from (SELECT ef.EXP_FORM_ID, fs.SUBJECT_NAME FROM FIN_EXPENSE_FORM ef " +
                    "LEFT JOIN FIN_EXPENSE_ITEM ei ON ef.EXP_FORM_ID = ei.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs ON ei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID)A " +
                    "WHERE ef.EXP_FORM_ID = A.EXP_FORM_ID FOR XML PATH('')) as Subjects FROM FIN_EXPENSE_FORM ef)B ";

                sql = sql + "WHERE B.STATUS = 40 AND ISNULL(B.PROJECT_ID, '') =@id ";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("id", projectid));
                // 費用發生年月條件
                if (null != occurreddate && occurreddate != "")
                {
                    sql = sql + "AND CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) =@occurreddate ";
                    parameters.Add(new SqlParameter("occurreddate", occurreddate));
                }
                //費用單編號條件
                if (null != expid && expid != "")
                {
                    sql = sql + "AND B.EXP_FORM_ID =@expid ";
                    parameters.Add(new SqlParameter("expid", expid));
                }
                //項目名稱條件
                if (null != subjectname && subjectname != "")
                {
                    sql = sql + "AND Subjects LIKE @subjectname ";
                    parameters.Add(new SqlParameter("subjectname", '%' + subjectname + '%'));
                }
                using (var context = new topmepEntities())
                {
                    logger.Debug("get expense form sql=" + sql);
                    lstForm = context.Database.SqlQuery<OperatingExpenseFunction>(sql, parameters.ToArray()).ToList();
                }
                logger.Info("get expense form count=" + lstForm.Count);
            }
            else
            {
                using (var context = new topmepEntities())
                {
                    lstForm = context.Database.SqlQuery<OperatingExpenseFunction>("SELECT B.EXP_FORM_ID, B.PAYMENT_DATE, B.STATUS, left(B.Subjects,len(B.Subjects)-1) AS SUBJECT_NAME, " +
                    "CONVERT(char(4), B.OCCURRED_YEAR) + '/' + CONVERT(char(2), B.OCCURRED_MONTH) AS OCCURRED_DATE, ROW_NUMBER() OVER(ORDER BY B.EXP_FORM_ID) AS NO " +
                    "FROM(SELECT ef.*,(SELECT cast( SUBJECT_NAME AS NVARCHAR ) + ',' from (SELECT ef.EXP_FORM_ID, fs.SUBJECT_NAME FROM FIN_EXPENSE_FORM ef " +
                    "LEFT JOIN FIN_EXPENSE_ITEM ei ON ef.EXP_FORM_ID = ei.EXP_FORM_ID LEFT JOIN FIN_SUBJECT fs ON ei.FIN_SUBJECT_ID = fs.FIN_SUBJECT_ID)A " +
                    "WHERE ef.EXP_FORM_ID = A.EXP_FORM_ID FOR XML PATH('')) as Subjects FROM FIN_EXPENSE_FORM ef)B WHERE B.STATUS < 20 AND ISNULL(B.PROJECT_ID, '') =@id ", new SqlParameter("id", projectid)).ToList();
                }
                logger.Info("get expense form count=" + lstForm.Count);
            }
            return lstForm;
        }

        //更新公司營業費用單狀態為退件
        public int RejectEXPByExpId(string expid)
        {
            int i = 0;
            logger.Info("reject EXP form by expid" + expid);
            string sql = "UPDATE  FIN_EXPENSE_FORM SET STATUS = 0 WHERE EXP_FORM_ID = @expid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("expid", expid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }

        //更新公司營業費用單狀態為主管已通過
        public int PassEXPByExpId(string expid, string passid)
        {
            int i = 0;
            logger.Info("Pass EXP form by expid" + expid);
            string sql = "UPDATE  FIN_EXPENSE_FORM SET STATUS = 30, PASS_CREATE_ID = @passid, PASS_CREATE_DATE = GETDATE() WHERE EXP_FORM_ID = @expid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("expid", expid));
            parameters.Add(new SqlParameter("passid", passid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }

        //更新公司營業費用單狀態為已立帳(即會計已完成稽核)
        public int JournalByExpId(string expid, string journalid)
        {
            int i = 0;
            logger.Info("Journal For EXP form by expid" + expid);
            string sql = "UPDATE  FIN_EXPENSE_FORM SET STATUS = 40, JOURNAL_CREATE_ID = @journalid, JOURNAL_CREATE_DATE = GETDATE() WHERE EXP_FORM_ID = @expid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("expid", expid));
            parameters.Add(new SqlParameter("journalid", journalid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }

        //更新公司營業費用單狀態為已核可
        public int ApproveEXPByExpId(string expid, string approveid)
        {
            int i = 0;
            logger.Info("Approve EXP form by expid" + expid);
            string sql = "UPDATE  FIN_EXPENSE_FORM SET STATUS = 50, APPROVE_CREATE_ID = @approveid, APPROVE_CREATE_DATE = GETDATE() WHERE EXP_FORM_ID = @expid ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("expid", expid));
            parameters.Add(new SqlParameter("approveid", approveid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }

        public int refreshAccountStatus(List<PLAN_ACCOUNT> lstItem)
        {
            int j = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("plan account item = " + lstItem.Count);
                //將item資料寫入 
                foreach (PLAN_ACCOUNT item in lstItem)
                {
                    PLAN_ACCOUNT existItem = null;
                    logger.Debug("form item id=" + item.PLAN_ACCOUNT_ID);
                    if (item.PLAN_ACCOUNT_ID != 0)
                    {
                        existItem = context.PLAN_ACCOUNT.Find(item.PLAN_ACCOUNT_ID);
                    }
                    else
                    {
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("itemid", item.ACCOUNT_FORM_ID));
                        string sql = "SELECT * FROM PLAN_ACCOUNT WHERE ACCOUNT_FORM_ID=@itemid";
                        logger.Info(sql + " ;" + item.ACCOUNT_FORM_ID);
                        PLAN_ACCOUNT excelItem = context.PLAN_ACCOUNT.SqlQuery(sql, parameters.ToArray()).First();
                        existItem = context.PLAN_ACCOUNT.Find(excelItem.PLAN_ACCOUNT_ID);

                    }
                    logger.Debug("find exist item=" + existItem.ACCOUNT_FORM_ID);
                    existItem.CHECK_NO = item.CHECK_NO;
                    existItem.PAYMENT_DATE = item.PAYMENT_DATE;
                    existItem.STATUS = item.STATUS;
                    context.PLAN_ACCOUNT.AddOrUpdate(existItem);
                }
                j = context.SaveChanges();
                logger.Debug("Update plan account item =" + j);
            }
            return j;
        }

        //取得符合條件之帳款資料
        public List<PlanAccountFunction> getPlanAccount(string paymentdate, string projectname, string payee, string accounttype, string formid, string duringStart, string duringEnd)
        {
            logger.Info("search plan account by " + paymentdate + ", 受款人 =" + payee + ", 專案名稱 =" + projectname + ", 帳款類型 =" + accounttype + ", 單據編號 =" + formid);
            List<PlanAccountFunction> lstForm = new List<PlanAccountFunction>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT ISNULL(pa.AMOUNT_PAYABLE, 0)AMOUNT_PAYABLE, ISNULL(pa.AMOUNT_PAID, 0)AMOUNT_PAID, pa.ACCOUNT_TYPE, CONVERT(char(10), pa.PAYMENT_DATE, 111) AS RECORDED_DATE, pa.PLAN_ACCOUNT_ID, " +
                "ISNULL(pa.STATUS, 10) AS STATUS , pa.CHECK_NO, p.PROJECT_NAME, pa.PAYEE, pa.REMARK, ROW_NUMBER() OVER(ORDER BY p.PROJECT_NAME) AS NO " +
                "FROM PLAN_ACCOUNT pa LEFT JOIN TND_PROJECT p ON pa.PROJECT_ID = p.PROJECT_ID WHERE pa.ACCOUNT_TYPE IN ('" + accounttype + "') ";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("accounttype", accounttype));

            //支付日期條件
            if (null != paymentdate && paymentdate != "")
            {
                sql = sql + "AND CONVERT(char(10), pa.PAYMENT_DATE, 111) =@paymentdate ";
                parameters.Add(new SqlParameter("paymentdate", paymentdate));
            }
            //專案名稱條件
            if (null != projectname && projectname != "")
            {
                sql = sql + "AND p.PROJECT_NAME LIKE @projectname ";
                parameters.Add(new SqlParameter("projectname", '%' + projectname + '%'));
            }
            //受款人條件
            if (null != payee && payee != "")
            {
                sql = sql + "AND pa.PAYEE LIKE @payee ";
                parameters.Add(new SqlParameter("payee", '%' + payee + '%'));
            }
            //單據編號條件
            if (null != formid && formid != "")
            {
                sql = sql + "AND pa.ACCOUNT_FORM_ID =@formid ";
                parameters.Add(new SqlParameter("formid", formid));
            }
            //支付區間條件
            if (null != duringStart && duringStart != "" || null != duringEnd && duringEnd != "")
            {
                sql = sql + "AND pa.PAYMENT_DATE >= convert(datetime, @duringStart, 111) AND pa.PAYMENT_DATE <= convert(datetime, @duringEnd, 111) ";
                parameters.Add(new SqlParameter("duringStart", duringStart));
                parameters.Add(new SqlParameter("duringEnd", duringEnd));
            }
            sql = sql + "ORDER BY p.PROJECT_NAME, pa.PAYMENT_DATE, pa.PAYEE DESC ";
            using (var context = new topmepEntities())
            {
                logger.Debug("get plan account sql=" + sql);
                lstForm = context.Database.SqlQuery<PlanAccountFunction>(sql, parameters.ToArray()).ToList();
                logger.Info("get plan account count=" + lstForm.Count);
            }
            return lstForm;
        }

        public PlanAccountFunction getPlanAccountItem(string itemid)
        {
            logger.Debug("get plan account item by id=" + itemid);
            PlanAccountFunction aitem = null;
            using (var context = new topmepEntities())
            {
                //條件篩選
                aitem = context.Database.SqlQuery<PlanAccountFunction>("SELECT PARSENAME(Convert(varchar,Convert(money,ISNULL(pa.AMOUNT_PAYABLE, 0)),1),2) AS RECORDED_AMOUNT_PAYABLE, " +
                    "PARSENAME(Convert(varchar,Convert(money,ISNULL(pa.AMOUNT_PAID, 0)),1),2) AS RECORDED_AMOUNT_PAID, " +
                    "pa.ACCOUNT_TYPE, CONVERT(char(10), pa.PAYMENT_DATE, 111) AS RECORDED_DATE, pa.REMARK, " +
                    "pa.PLAN_ACCOUNT_ID, pa.CONTRACT_ID, pa.ACCOUNT_TYPE, pa.ACCOUNT_FORM_ID, pa.ISDEBIT, ISNULL(pa.STATUS, 10) AS STATUS, " +
                    "pa.CREATE_ID, pa.PROJECT_ID, pa.CHECK_NO, p.PROJECT_NAME FROM PLAN_ACCOUNT pa LEFT JOIN TND_PROJECT p ON pa.PROJECT_ID = p.PROJECT_ID " +
                    "WHERE pa.PLAN_ACCOUNT_ID=@itemid ",
                new SqlParameter("itemid", itemid)).First();
            }
            return aitem;
        }
        //將Plan Account Item 刪除
        public int delPlanAccountItem(string itemid)
        {
            string sql = "DELETE FROM PLAN_ACCOUNT WHERE PLAN_ACCOUNT_ID = @itemid";
            int i = 0;
            using (var db = new topmepEntities())
            {
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("itemid", itemid));
                logger.Info("Delete PLAN_ACCOUNT =" + sql + ",itemid=" + itemid);
                i = db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            }
            return i;
        }
        //取得特定日期借款與還款
        public List<LoanTranactionFunction> getLoanTranaction(string type, string paymentdate, string duringStart, string duringEnd)
        {
            logger.Debug("get loan tranaction by payment date=" + paymentdate + ",and type = " + type);
            List<LoanTranactionFunction> lstItem = new List<LoanTranactionFunction>();
            using (var context = new topmepEntities())
            {
                //條件篩選
                if (null != paymentdate && paymentdate != "")
                {
                    if (type == "I")// type值為'I'表示有現金流入
                    {
                        lstItem = context.Database.SqlQuery<LoanTranactionFunction>("SELECT t.*, l.IS_SUPPLIER FROM FIN_LOAN_TRANACTION t LEFT JOIN FIN_BANK_LOAN l ON t.BL_ID = l.BL_ID  WHERE ISNULL(l.IS_SUPPLIER, 'N') = 'Y' AND t.TRANSACTION_TYPE = 1 AND CONVERT(char(10), " +
                            "IIF(TRANSACTION_TYPE = 1,PAYBACK_DATE,EVENT_DATE), 111) =@paymentdate OR ISNULL(l.IS_SUPPLIER, 'N') <> 'Y' AND t.REMARK NOT LIKE '%備償%' AND t.TRANSACTION_TYPE = -1 AND CONVERT(char(10), IIF(TRANSACTION_TYPE = 1, PAYBACK_DATE, EVENT_DATE), 111) =@paymentdate ",
                        new SqlParameter("paymentdate", paymentdate)).ToList();
                    }
                    else // type值為'O'表示有現金流出
                    {
                        lstItem = context.Database.SqlQuery<LoanTranactionFunction>("SELECT t.*, l.IS_SUPPLIER FROM FIN_LOAN_TRANACTION t LEFT JOIN FIN_BANK_LOAN l ON t.BL_ID = l.BL_ID  WHERE ISNULL(l.IS_SUPPLIER, 'N') = 'Y' AND t.TRANSACTION_TYPE = -1 AND CONVERT(char(10), " +
                            "IIF(TRANSACTION_TYPE = 1,PAYBACK_DATE,EVENT_DATE), 111) =@paymentdate OR ISNULL(l.IS_SUPPLIER, 'N') <> 'Y' AND t.REMARK NOT LIKE '%備償%' AND t.TRANSACTION_TYPE = 1 AND CONVERT(char(10), IIF(TRANSACTION_TYPE = 1, PAYBACK_DATE, EVENT_DATE), 111) =@paymentdate ",
                        new SqlParameter("paymentdate", paymentdate)).ToList();
                    }
                }
                else if (null != duringStart && duringStart != "" && null != duringEnd && duringEnd != "")
                {
                    if (type == "I")// type值為'I'表示有現金流入
                    {
                        lstItem = context.Database.SqlQuery<LoanTranactionFunction>("SELECT t.*, l.IS_SUPPLIER FROM FIN_LOAN_TRANACTION t LEFT JOIN FIN_BANK_LOAN l ON t.BL_ID = l.BL_ID  WHERE ISNULL(l.IS_SUPPLIER, 'N') = 'Y' AND t.TRANSACTION_TYPE = 1 AND " +
                            "IIF(TRANSACTION_TYPE = 1,PAYBACK_DATE,EVENT_DATE) >= convert(datetime, @duringStart, 111) AND IIF(TRANSACTION_TYPE = 1,PAYBACK_DATE,EVENT_DATE) <= convert(datetime, @duringEnd, 111) OR ISNULL(l.IS_SUPPLIER, 'N') <> 'Y' " +
                            "AND t.REMARK NOT LIKE '%備償%' AND t.TRANSACTION_TYPE = -1 AND IIF(TRANSACTION_TYPE = 1,PAYBACK_DATE,EVENT_DATE) >= convert(datetime, @duringStart, 111) AND IIF(TRANSACTION_TYPE = 1,PAYBACK_DATE,EVENT_DATE) <= convert(datetime, @duringEnd, 111) ORDER BY IIF(TRANSACTION_TYPE = 1,PAYBACK_DATE,EVENT_DATE) ",
                        new SqlParameter("duringStart", duringStart), new SqlParameter("duringEnd", duringEnd)).ToList();
                    }
                    else // type值為'O'表示有現金流出
                    {
                        lstItem = context.Database.SqlQuery<LoanTranactionFunction>("SELECT t.*, l.IS_SUPPLIER FROM FIN_LOAN_TRANACTION t LEFT JOIN FIN_BANK_LOAN l ON t.BL_ID = l.BL_ID  WHERE ISNULL(l.IS_SUPPLIER, 'N') = 'Y' AND t.TRANSACTION_TYPE = -1 AND " +
                            "IIF(TRANSACTION_TYPE = 1,PAYBACK_DATE,EVENT_DATE) >= convert(datetime, @duringStart, 111) AND IIF(TRANSACTION_TYPE = 1,PAYBACK_DATE,EVENT_DATE) <= convert(datetime, @duringEnd, 111) OR ISNULL(l.IS_SUPPLIER, 'N') <> 'Y' " +
                            "AND t.REMARK NOT LIKE '%備償%' AND t.TRANSACTION_TYPE = 1 AND IIF(TRANSACTION_TYPE = 1,PAYBACK_DATE,EVENT_DATE) >= convert(datetime, @duringStart, 111) AND IIF(TRANSACTION_TYPE = 1,PAYBACK_DATE,EVENT_DATE) <= convert(datetime, @duringEnd, 111) ORDER BY IIF(TRANSACTION_TYPE = 1,PAYBACK_DATE,EVENT_DATE) ",
                       new SqlParameter("duringStart", duringStart), new SqlParameter("duringEnd", duringEnd)).ToList();
                    }
                }
                else
                {
                    if (type == "I")// type值為'I'表示有現金流入
                    {
                        lstItem = context.Database.SqlQuery<LoanTranactionFunction>("SELECT t.*, l.IS_SUPPLIER FROM FIN_LOAN_TRANACTION t LEFT JOIN FIN_BANK_LOAN l ON t.BL_ID = l.BL_ID  WHERE ISNULL(l.IS_SUPPLIER, 'N') = 'Y' AND t.TRANSACTION_TYPE = 1 " +
                            "OR ISNULL(l.IS_SUPPLIER, 'N') <> 'Y' AND t.REMARK NOT LIKE '%備償%' AND t.TRANSACTION_TYPE = -1 ORDER BY IIF(TRANSACTION_TYPE = 1,PAYBACK_DATE,EVENT_DATE) ",
                        new SqlParameter("paymentdate", paymentdate)).ToList();
                    }
                    else // type值為'O'表示有現金流出
                    {
                        lstItem = context.Database.SqlQuery<LoanTranactionFunction>("SELECT t.*, l.IS_SUPPLIER FROM FIN_LOAN_TRANACTION t LEFT JOIN FIN_BANK_LOAN l ON t.BL_ID = l.BL_ID  WHERE ISNULL(l.IS_SUPPLIER, 'N') = 'Y' AND t.TRANSACTION_TYPE = -1 " +
                            "OR ISNULL(l.IS_SUPPLIER, 'N') <> 'Y' AND t.REMARK NOT LIKE '%備償%' AND t.TRANSACTION_TYPE = 1 ORDER BY IIF(TRANSACTION_TYPE = 1,PAYBACK_DATE,EVENT_DATE) ",
                        new SqlParameter("paymentdate", paymentdate)).ToList();
                    }
                }
            }
            return lstItem;
        }

        public List<PlanAccountFunction> getOutFlowBalanceByDate(string paymentDate, string duringStart, string duringEnd)
        {
            logger.Debug("get cash out flow balance by payment date, and payment date =" + paymentDate);
            List<PlanAccountFunction> lstItem = new List<PlanAccountFunction>();
            using (var context = new topmepEntities())
            {
                //條件篩選
                if (null != paymentDate && paymentDate != "")
                {
                    lstItem = context.Database.SqlQuery<PlanAccountFunction>("SELECT p.* , CONVERT(char(10), p.PAYMENT_DATE, 111) AS RECORDED_DATE, PARSENAME(Convert(varchar,Convert(money,ISNULL(p.AMOUNT_PAID, 0)),1),2) AS RECORDED_AMOUNT_PAYABLE, " +
                           "PARSENAME(Convert(varchar, Convert(money, ISNULL(it.AMOUNT, 0)), 1), 2) AS PAYBACK_AMOUNT, PARSENAME(Convert(varchar, Convert(money, ISNULL(p.AMOUNT_PAID, 0) - ISNULL(it.AMOUNT, 0)), 1), 2) AS RECORDED_AMOUNT_PAID  FROM " +
                           "(select o.PAYEE, CONVERT(datetime,o.PAYMENT_DATE, 111)PAYMENT_DATE, SUM(o.AMOUNT)AMOUNT_PAID from(SELECT CONVERT(varchar, PAYMENT_DATE, 111)PAYMENT_DATE, SUM(fei.AMOUNT)AMOUNT, PAYEE FROM FIN_EXPENSE_FORM fef LEFT JOIN FIN_EXPENSE_ITEM fei ON fef.EXP_FORM_ID = fei.EXP_FORM_ID " +
                           "WHERE fef.STATUS = 30 AND CONVERT(varchar, PAYMENT_DATE, 111) = @paymentDate GROUP BY CONVERT(varchar, PAYMENT_DATE, 111), PAYEE " +
                           "union SELECT CONVERT(char(10), PAYMENT_DATE, 111)PAYMENT_DATE, SUM(AMOUNT_PAID)AMOUNT_PAID, PAYEE FROM PLAN_ACCOUNT WHERE ISDEBIT = 'N' AND CONVERT(char(10), PAYMENT_DATE, 111) = @paymentDate GROUP BY PAYEE, CONVERT(char(10), PAYMENT_DATE, 111))o GROUP BY o.PAYEE, o.PAYMENT_DATE)p " +
                           "LEFT JOIN(SELECT SUM(t.AMOUNT)AMOUNT, l.BANK_NAME FROM FIN_LOAN_TRANACTION t LEFT JOIN FIN_BANK_LOAN l ON t.BL_ID = l.BL_ID  WHERE ISNULL(l.IS_SUPPLIER, 'N') = 'Y' AND t.TRANSACTION_TYPE = 1 AND CONVERT(char(10),IIF(TRANSACTION_TYPE = 1, PAYBACK_DATE, EVENT_DATE), 111) = @paymentDate " +
                           "OR ISNULL(l.IS_SUPPLIER, 'N') <> 'Y' AND t.REMARK NOT LIKE '%備償%' AND t.TRANSACTION_TYPE = -1 AND CONVERT(char(10), IIF(TRANSACTION_TYPE = 1, PAYBACK_DATE, EVENT_DATE), 111) = @paymentDate GROUP BY l.BANK_NAME)it " +
                           "ON it.BANK_NAME = p.PAYEE ", new SqlParameter("paymentDate", paymentDate)).ToList();
                }
                else if (null != duringStart && duringStart != "" && null != duringEnd && duringEnd != "")
                {
                    lstItem = context.Database.SqlQuery<PlanAccountFunction>("SELECT p.* , CONVERT(char(10), p.PAYMENT_DATE, 111) AS RECORDED_DATE, PARSENAME(Convert(varchar,Convert(money,ISNULL(p.AMOUNT_PAID, 0)),1),2) AS RECORDED_AMOUNT_PAYABLE, " +
                          "PARSENAME(Convert(varchar, Convert(money, ISNULL(it.AMOUNT, 0)), 1), 2) AS PAYBACK_AMOUNT, PARSENAME(Convert(varchar, Convert(money, ISNULL(p.AMOUNT_PAID, 0) - ISNULL(it.AMOUNT, 0)), 1), 2) AS RECORDED_AMOUNT_PAID  FROM " +
                          "(select o.PAYEE, CONVERT(datetime,o.PAYMENT_DATE, 111)PAYMENT_DATE, SUM(o.AMOUNT)AMOUNT_PAID from(SELECT CONVERT(varchar, PAYMENT_DATE, 111)PAYMENT_DATE, SUM(fei.AMOUNT)AMOUNT, PAYEE FROM FIN_EXPENSE_FORM fef LEFT JOIN FIN_EXPENSE_ITEM fei ON fef.EXP_FORM_ID = fei.EXP_FORM_ID " +
                          "WHERE fef.STATUS = 30 AND PAYMENT_DATE >= CONVERT(datetime, @duringStart, 111) AND PAYMENT_DATE <= CONVERT(varchar, @duringEnd, 111) GROUP BY CONVERT(varchar, PAYMENT_DATE, 111), PAYEE " +
                          "union SELECT CONVERT(char(10), PAYMENT_DATE, 111)PAYMENT_DATE, SUM(AMOUNT_PAID)AMOUNT_PAID, PAYEE FROM PLAN_ACCOUNT WHERE ISDEBIT = 'N' AND PAYMENT_DATE >= CONVERT(datetime, @duringStart, 111) AND " +
                          "PAYMENT_DATE <= CONVERT(datetime, @duringEnd, 111) GROUP BY PAYEE, CONVERT(char(10), PAYMENT_DATE, 111))o GROUP BY o.PAYEE, o.PAYMENT_DATE)p " +
                          "LEFT JOIN(SELECT SUM(t.AMOUNT)AMOUNT, l.BANK_NAME FROM FIN_LOAN_TRANACTION t LEFT JOIN FIN_BANK_LOAN l ON t.BL_ID = l.BL_ID  WHERE ISNULL(l.IS_SUPPLIER, 'N') = 'Y' AND t.TRANSACTION_TYPE = 1 AND " +
                          "IIF(TRANSACTION_TYPE = 1, PAYBACK_DATE, EVENT_DATE) >= CONVERT(datetime, @duringStart, 111) AND IIF(TRANSACTION_TYPE = 1, PAYBACK_DATE, EVENT_DATE) <= CONVERT(datetime, @duringEnd, 111) OR " +
                          "ISNULL(l.IS_SUPPLIER, 'N') <> 'Y' AND t.REMARK NOT LIKE '%備償%' AND t.TRANSACTION_TYPE = -1 AND IIF(TRANSACTION_TYPE = 1, PAYBACK_DATE, EVENT_DATE) >= CONVERT(datetime, @duringStart, 111) AND " +
                          "IIF(TRANSACTION_TYPE = 1, PAYBACK_DATE, EVENT_DATE) <= CONVERT(datetime, @duringEnd, 111) GROUP BY l.BANK_NAME)it " +
                          "ON it.BANK_NAME = p.PAYEE ORDER BY p.PAYMENT_DATE ", new SqlParameter("duringStart", duringStart), new SqlParameter("duringEnd", duringEnd)).ToList();
                }
                else
                {
                    lstItem = context.Database.SqlQuery<PlanAccountFunction>("SELECT p.* , CONVERT(char(10), p.PAYMENT_DATE, 111) AS RECORDED_DATE, PARSENAME(Convert(varchar,Convert(money,ISNULL(p.AMOUNT_PAID, 0)),1),2) AS RECORDED_AMOUNT_PAYABLE, " +
                           "PARSENAME(Convert(varchar, Convert(money, ISNULL(it.AMOUNT, 0)), 1), 2) AS PAYBACK_AMOUNT, PARSENAME(Convert(varchar, Convert(money, ISNULL(p.AMOUNT_PAID, 0) - ISNULL(it.AMOUNT, 0)), 1), 2) AS RECORDED_AMOUNT_PAID  FROM " +
                           "(select o.PAYEE, CONVERT(datetime,o.PAYMENT_DATE, 111)PAYMENT_DATE, SUM(o.AMOUNT)AMOUNT_PAID from(SELECT CONVERT(varchar, PAYMENT_DATE, 111)PAYMENT_DATE, SUM(fei.AMOUNT)AMOUNT, PAYEE FROM FIN_EXPENSE_FORM fef LEFT JOIN FIN_EXPENSE_ITEM fei ON fef.EXP_FORM_ID = fei.EXP_FORM_ID " +
                           "WHERE fef.STATUS = 30 GROUP BY CONVERT(varchar, PAYMENT_DATE, 111), PAYEE " +
                           "union SELECT CONVERT(char(10), PAYMENT_DATE, 111)PAYMENT_DATE, SUM(AMOUNT_PAID)AMOUNT_PAID, PAYEE FROM PLAN_ACCOUNT WHERE ISDEBIT = 'N' GROUP BY PAYEE, CONVERT(char(10), PAYMENT_DATE, 111))o GROUP BY o.PAYEE, o.PAYMENT_DATE)p " +
                           "LEFT JOIN(SELECT SUM(t.AMOUNT)AMOUNT, l.BANK_NAME FROM FIN_LOAN_TRANACTION t LEFT JOIN FIN_BANK_LOAN l ON t.BL_ID = l.BL_ID  WHERE ISNULL(l.IS_SUPPLIER, 'N') = 'Y' AND t.TRANSACTION_TYPE = 1 " +
                           "OR ISNULL(l.IS_SUPPLIER, 'N') <> 'Y' AND t.REMARK NOT LIKE '%備償%' AND t.TRANSACTION_TYPE = -1 GROUP BY l.BANK_NAME)it " +
                           "ON it.BANK_NAME = p.PAYEE ORDER BY p.PAYMENT_DATE ").ToList();
                }
            }
            return lstItem;
        }

        public List<ExpenseFormFunction> getExpenseOutFlowByDate(string paymentDate, string duringStart, string duringEnd)
        {
            logger.Debug("get cash out flow from expense form by payment date, and payment date =" + paymentDate);
            List<ExpenseFormFunction> lstItem = new List<ExpenseFormFunction>();
            using (var context = new topmepEntities())
            {
                //條件篩選
                if (null != paymentDate && paymentDate != "")
                {
                    lstItem = context.Database.SqlQuery<ExpenseFormFunction>("SELECT fef.EXP_FORM_ID, ISNULL(fef.PROJECT_ID, '')PROJECT_ID, fef.PAYEE, CONVERT(varchar, fef.PAYMENT_DATE, 111)RECORDED_DATE, SUM(fei.AMOUNT)AMOUNT " +
                        "FROM FIN_EXPENSE_FORM fef LEFT JOIN FIN_EXPENSE_ITEM fei ON fef.EXP_FORM_ID = fei.EXP_FORM_ID WHERE fef.STATUS = 30 AND CONVERT(varchar, fef.PAYMENT_DATE, 111) = @paymentDate " +
                        "GROUP BY fef.EXP_FORM_ID, fef.PAYEE, ISNULL(fef.PROJECT_ID, ''), CONVERT(varchar, fef.PAYMENT_DATE, 111) ", new SqlParameter("paymentDate", paymentDate)).ToList();
                }
                else if (null != duringStart && duringStart != "" && null != duringEnd && duringEnd != "")
                {
                    lstItem = context.Database.SqlQuery<ExpenseFormFunction>("SELECT fef.EXP_FORM_ID, ISNULL(fef.PROJECT_ID, '')PROJECT_ID, fef.PAYEE, CONVERT(varchar, fef.PAYMENT_DATE, 111)RECORDED_DATE, SUM(fei.AMOUNT)AMOUNT " +
                        "FROM FIN_EXPENSE_FORM fef LEFT JOIN FIN_EXPENSE_ITEM fei ON fef.EXP_FORM_ID = fei.EXP_FORM_ID WHERE fef.STATUS = 30 AND CONVERT(varchar, fef.PAYMENT_DATE, 111) >= @duringStart " +
                        "AND CONVERT(varchar, fef.PAYMENT_DATE, 111) <= @duringEnd GROUP BY fef.EXP_FORM_ID, fef.PAYEE, ISNULL(fef.PROJECT_ID, ''), CONVERT(varchar, fef.PAYMENT_DATE, 111) "
                        , new SqlParameter("duringStart", duringStart), new SqlParameter("duringEnd", duringEnd)).ToList();
                }
                else
                {
                    lstItem = context.Database.SqlQuery<ExpenseFormFunction>("SELECT fef.EXP_FORM_ID, ISNULL(fef.PROJECT_ID, '')PROJECT_ID, fef.PAYEE, CONVERT(varchar, fef.PAYMENT_DATE, 111)RECORDED_DATE, SUM(fei.AMOUNT)AMOUNT " +
                        "FROM FIN_EXPENSE_FORM fef LEFT JOIN FIN_EXPENSE_ITEM fei ON fef.EXP_FORM_ID = fei.EXP_FORM_ID WHERE fef.STATUS = 30 GROUP BY fef.EXP_FORM_ID, fef.PAYEE, ISNULL(fef.PROJECT_ID, ''), CONVERT(varchar, fef.PAYMENT_DATE, 111) ").ToList();
                }
            }
            return lstItem;
        }

        public List<ExpenseFormFunction> getExpenseBudgetByDate(string paymentDate, string duringStart, string duringEnd)
        {
            logger.Debug("get expense budget by payment date, and payment date =" + paymentDate);
            List<ExpenseFormFunction> lstItem = new List<ExpenseFormFunction>();
            using (var context = new topmepEntities())
            {
                //條件篩選
                if (null != paymentDate && paymentDate != "")
                {
                    string sql = @"SELECT CONVERT(char(10),CONVERT(datetime, exp.PaidDate , 111), 111)RECORDED_DATE, exp.AMOUNT 
                            from (SELECT CONVERT(varchar, r.CURRENT_YEAR) + '/'+ CONVERT(varchar, r.BUDGET_MONTH) + '/'+ CONVERT(varchar, 12)PaidDate, sum(PaidAmt)AMOUNT 
                            from (SELECT budget.*, iif(expense.AMOUNT_REAL is not null, 0, budget.AMOUNT)PaidAmt 
                            FROM (SELECT SUBJECT_ID, CURRENT_YEAR, BUDGET_MONTH, AMOUNT 
                            FROM FIN_EXPENSE_BUDGET WHERE SUBJECT_ID = '6020' AND CURRENT_YEAR >= YEAR(GETDATE()) 
                            UNION 
                            SELECT SUBJECT_ID, BUDGET_YEAR, BUDGET_MONTH, AMOUNT 
                            FROM PLAN_SITE_BUDGET WHERE SUBJECT_ID = '1312' AND  BUDGET_YEAR >= YEAR(GETDATE()))budget left join(SELECT OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID, sum(i.AMOUNT) as AMOUNT_REAL 
                            FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM i ON f.EXP_FORM_ID = i.EXP_FORM_ID 
                            WHERE i.FIN_SUBJECT_ID in ('6020', '1312') AND f.STATUS = 30 GROUP BY OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID)expense 
                            ON budget.SUBJECT_ID + CONVERT(varchar, budget.CURRENT_YEAR) + CONVERT(varchar, budget.BUDGET_MONTH) = expense.FIN_SUBJECT_ID + CONVERT(varchar, expense.OCCURRED_YEAR) + CONVERT(varchar, expense.OCCURRED_MONTH))r GROUP BY r.CURRENT_YEAR,r.BUDGET_MONTH 
                            UNION 
                            SELECT CONVERT(varchar, r.CURRENT_YEAR) +'/' + CONVERT(varchar, r.BUDGET_MONTH) + '/' + CONVERT(varchar, 5)PaidDate, sum(PaidAmt)AMOUNT from(SELECT budget.*, iif(expense.AMOUNT_REAL is not null, 0, budget.AMOUNT)PaidAmt 
                            FROM(SELECT SUBJECT_ID, CURRENT_YEAR, BUDGET_MONTH, AMOUNT FROM FIN_EXPENSE_BUDGET WHERE SUBJECT_ID = '6010' AND CURRENT_YEAR >= YEAR(GETDATE()) UNION 
                            SELECT SUBJECT_ID, BUDGET_YEAR, BUDGET_MONTH, AMOUNT FROM PLAN_SITE_BUDGET WHERE SUBJECT_ID = '1301' AND  BUDGET_YEAR >= YEAR(GETDATE()))budget 
                            left join(SELECT OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID, sum(i.AMOUNT) as AMOUNT_REAL 
                            FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM i ON f.EXP_FORM_ID = i.EXP_FORM_ID 
                            WHERE i.FIN_SUBJECT_ID in ('6010', '1301') AND f.STATUS = 30 GROUP BY OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID)expense ON budget.SUBJECT_ID + CONVERT(varchar, budget.CURRENT_YEAR) + CONVERT(varchar, budget.BUDGET_MONTH) = expense.FIN_SUBJECT_ID + CONVERT(varchar, expense.OCCURRED_YEAR) + CONVERT(varchar, expense.OCCURRED_MONTH))r GROUP BY r.CURRENT_YEAR,r.BUDGET_MONTH 
                            UNION 
                            SELECT CONVERT(varchar, r.CURRENT_YEAR) +'/' + CONVERT(varchar, r.BUDGET_MONTH) + '/' + CONVERT(varchar, 15)PaidDate, sum(PaidAmt)AMOUNT 
                            from(SELECT budget.*, iif(expense.AMOUNT_REAL is not null, 0, budget.AMOUNT)PaidAmt 
                            FROM(SELECT SUBJECT_ID, CURRENT_YEAR, BUDGET_MONTH, AMOUNT FROM FIN_EXPENSE_BUDGET WHERE SUBJECT_ID = '6100' AND CURRENT_YEAR >= YEAR(GETDATE()) 
                            UNION 
                            SELECT SUBJECT_ID, BUDGET_YEAR, BUDGET_MONTH, AMOUNT FROM PLAN_SITE_BUDGET WHERE SUBJECT_ID = '1306' AND  BUDGET_YEAR >= YEAR(GETDATE()))budget left join(SELECT OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID, sum(i.AMOUNT) as AMOUNT_REAL 
                            FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM i ON f.EXP_FORM_ID = i.EXP_FORM_ID 
                            WHERE i.FIN_SUBJECT_ID in ('6100', '1306') AND f.STATUS = 30 GROUP BY OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID)expense ON budget.SUBJECT_ID + CONVERT(varchar, budget.CURRENT_YEAR) + CONVERT(varchar, budget.BUDGET_MONTH) = expense.FIN_SUBJECT_ID + CONVERT(varchar, expense.OCCURRED_YEAR) + CONVERT(varchar, expense.OCCURRED_MONTH))r GROUP BY r.CURRENT_YEAR,r.BUDGET_MONTH 
                            UNION 
                            SELECT CONVERT(varchar, r.CURRENT_YEAR) +'/' + CONVERT(varchar, r.BUDGET_MONTH) + '/' + iif(BUDGET_MONTH <> 2, CONVERT(varchar, 30), CONVERT(varchar, 28))PaidDate, sum(PaidAmt)PaidAmt from(SELECT budget.*, iif(expense.AMOUNT_REAL is not null, 0, budget.AMOUNT)PaidAmt FROM(SELECT SUBJECT_ID, CURRENT_YEAR, BUDGET_MONTH, AMOUNT 
                            FROM FIN_EXPENSE_BUDGET WHERE SUBJECT_ID = '6560' AND CURRENT_YEAR >= YEAR(GETDATE()) 
                            UNION 
                            SELECT SUBJECT_ID, BUDGET_YEAR, BUDGET_MONTH, AMOUNT FROM PLAN_SITE_BUDGET WHERE SUBJECT_ID = '1307' AND  BUDGET_YEAR >= YEAR(GETDATE()))budget left join(SELECT OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID, sum(i.AMOUNT) as AMOUNT_REAL 
                            FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM i ON f.EXP_FORM_ID = i.EXP_FORM_ID 
                            WHERE i.FIN_SUBJECT_ID in ('6560', '1307') AND f.STATUS = 30 GROUP BY OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID)expense ON budget.SUBJECT_ID + CONVERT(varchar, budget.CURRENT_YEAR) + CONVERT(varchar, budget.BUDGET_MONTH) = expense.FIN_SUBJECT_ID + CONVERT(varchar, expense.OCCURRED_YEAR) + CONVERT(varchar, expense.OCCURRED_MONTH))r GROUP BY r.CURRENT_YEAR,r.BUDGET_MONTH 
                            UNION SELECT CONVERT(varchar, r.CURRENT_YEAR) +'/' + CONVERT(varchar, r.BUDGET_MONTH) + '/' + CONVERT(varchar, 5)PaidDate, sum(PaidAmt)AMOUNT from(SELECT budget.*, iif(expense.AMOUNT_REAL is not null, 0, budget.AMOUNT)PaidAmt FROM(SELECT SUBJECT_ID, BUDGET_YEAR AS CURRENT_YEAR, BUDGET_MONTH, AMOUNT FROM PLAN_SITE_BUDGET 
                            WHERE SUBJECT_ID IN('1321', '1317', '1318', '1316', '1328', '1315', '1319') AND BUDGET_YEAR >= YEAR(GETDATE()))budget left join(SELECT OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID, sum(i.AMOUNT) as AMOUNT_REAL FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM i ON f.EXP_FORM_ID = i.EXP_FORM_ID 
                            WHERE i.FIN_SUBJECT_ID in ('1321', '1317', '1318', '1316', '1328', '1315', '1319') AND f.STATUS = 30 GROUP BY OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID)expense ON budget.SUBJECT_ID + CONVERT(varchar, budget.CURRENT_YEAR) + CONVERT(varchar, budget.BUDGET_MONTH) = expense.FIN_SUBJECT_ID + CONVERT(varchar, expense.OCCURRED_YEAR) + CONVERT(varchar, expense.OCCURRED_MONTH))r GROUP BY r.CURRENT_YEAR,r.BUDGET_MONTH)exp
                            where exp.AMOUNT <> 0 and CONVERT(char(10), CONVERT(datetime, exp.PaidDate, 111), 111) = @paymentDate and CONVERT(char(10), CONVERT(datetime, exp.PaidDate, 111), 111) >= CONVERT(char(10), GETDATE(), 111) ";
                    logger.Debug("sql=" + sql + ",paymentdate=" + paymentDate);
                    lstItem = context.Database.SqlQuery<ExpenseFormFunction>(sql, new SqlParameter("paymentDate", paymentDate)).ToList();
                }
                else if (null != duringStart && duringStart != "" && null != duringEnd && duringEnd != "")
                {
                    string sql = @"select CONVERT(char(10),CONVERT(datetime, exp.PaidDate , 111), 111)RECORDED_DATE, exp.AMOUNT from (SELECT CONVERT(varchar, r.CURRENT_YEAR) + '/'+ CONVERT(varchar, r.BUDGET_MONTH) + '/'+ CONVERT(varchar, 12)PaidDate, sum(PaidAmt)AMOUNT from (SELECT budget.*, iif(expense.AMOUNT_REAL is not null, 0, budget.AMOUNT)PaidAmt 
                        FROM (SELECT SUBJECT_ID, CURRENT_YEAR, BUDGET_MONTH, AMOUNT FROM FIN_EXPENSE_BUDGET WHERE SUBJECT_ID = '6020' AND CURRENT_YEAR >= YEAR(GETDATE()) 
                        UNION 
                        SELECT SUBJECT_ID, BUDGET_YEAR, BUDGET_MONTH, AMOUNT FROM PLAN_SITE_BUDGET WHERE SUBJECT_ID = '1312' AND  BUDGET_YEAR >= YEAR(GETDATE()))budget left join(SELECT OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID, sum(i.AMOUNT) as AMOUNT_REAL 
                        FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM i ON f.EXP_FORM_ID = i.EXP_FORM_ID WHERE i.FIN_SUBJECT_ID in ('6020', '1312') AND f.STATUS = 30 GROUP BY OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID)expense 
                        ON budget.SUBJECT_ID + CONVERT(varchar, budget.CURRENT_YEAR) + CONVERT(varchar, budget.BUDGET_MONTH) = expense.FIN_SUBJECT_ID + CONVERT(varchar, expense.OCCURRED_YEAR) + CONVERT(varchar, expense.OCCURRED_MONTH))r GROUP BY r.CURRENT_YEAR,r.BUDGET_MONTH 
                        UNION 
                        SELECT CONVERT(varchar, r.CURRENT_YEAR) +'/' + CONVERT(varchar, r.BUDGET_MONTH) + '/' + CONVERT(varchar, 5)PaidDate, sum(PaidAmt)AMOUNT from(SELECT budget.*, iif(expense.AMOUNT_REAL is not null, 0, budget.AMOUNT)PaidAmt FROM(SELECT SUBJECT_ID, CURRENT_YEAR, BUDGET_MONTH, AMOUNT FROM FIN_EXPENSE_BUDGET WHERE SUBJECT_ID = '6010' AND CURRENT_YEAR >= YEAR(GETDATE()) UNION 
                        SELECT SUBJECT_ID, BUDGET_YEAR, BUDGET_MONTH, AMOUNT FROM PLAN_SITE_BUDGET WHERE SUBJECT_ID = '1301' AND  BUDGET_YEAR >= YEAR(GETDATE()))budget left join(SELECT OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID, sum(i.AMOUNT) as AMOUNT_REAL FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM i ON f.EXP_FORM_ID = i.EXP_FORM_ID 
                        WHERE i.FIN_SUBJECT_ID in ('6010', '1301') AND f.STATUS = 30 GROUP BY OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID)expense ON budget.SUBJECT_ID + CONVERT(varchar, budget.CURRENT_YEAR) + CONVERT(varchar, budget.BUDGET_MONTH) = expense.FIN_SUBJECT_ID + CONVERT(varchar, expense.OCCURRED_YEAR) + CONVERT(varchar, expense.OCCURRED_MONTH))r GROUP BY r.CURRENT_YEAR,r.BUDGET_MONTH
                        UNION
                        SELECT CONVERT(varchar, r.CURRENT_YEAR) +'/' + CONVERT(varchar, r.BUDGET_MONTH) + '/' + CONVERT(varchar, 15)PaidDate, sum(PaidAmt)AMOUNT from(SELECT budget.*, iif(expense.AMOUNT_REAL is not null, 0, budget.AMOUNT)PaidAmt FROM(SELECT SUBJECT_ID, CURRENT_YEAR, BUDGET_MONTH, AMOUNT FROM FIN_EXPENSE_BUDGET WHERE SUBJECT_ID = '6100' AND CURRENT_YEAR >= YEAR(GETDATE()) UNION
                        SELECT SUBJECT_ID, BUDGET_YEAR, BUDGET_MONTH, AMOUNT FROM PLAN_SITE_BUDGET WHERE SUBJECT_ID = '1306' AND  BUDGET_YEAR >= YEAR(GETDATE()))budget left join(SELECT OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID, sum(i.AMOUNT) as AMOUNT_REAL FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM i ON f.EXP_FORM_ID = i.EXP_FORM_ID 
                        WHERE i.FIN_SUBJECT_ID in ('6100', '1306') AND f.STATUS = 30 GROUP BY OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID)expense ON budget.SUBJECT_ID + CONVERT(varchar, budget.CURRENT_YEAR) + CONVERT(varchar, budget.BUDGET_MONTH) = expense.FIN_SUBJECT_ID + CONVERT(varchar, expense.OCCURRED_YEAR) + CONVERT(varchar, expense.OCCURRED_MONTH))r GROUP BY r.CURRENT_YEAR,r.BUDGET_MONTH
                        UNION
                        SELECT CONVERT(varchar, r.CURRENT_YEAR) +'/' + CONVERT(varchar, r.BUDGET_MONTH) + '/' + iif(BUDGET_MONTH <> 2, CONVERT(varchar, 30), CONVERT(varchar, 28))PaidDate, sum(PaidAmt)PaidAmt from(SELECT budget.*, iif(expense.AMOUNT_REAL is not null, 0, budget.AMOUNT)PaidAmt FROM(SELECT SUBJECT_ID, CURRENT_YEAR, BUDGET_MONTH, AMOUNT FROM FIN_EXPENSE_BUDGET WHERE SUBJECT_ID = '6560'
                        AND CURRENT_YEAR >= YEAR(GETDATE()) UNION 
                        SELECT SUBJECT_ID, BUDGET_YEAR, BUDGET_MONTH, AMOUNT FROM PLAN_SITE_BUDGET WHERE SUBJECT_ID = '1307' AND  BUDGET_YEAR >= YEAR(GETDATE()))budget left join(SELECT OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID, sum(i.AMOUNT) as AMOUNT_REAL FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM i ON f.EXP_FORM_ID = i.EXP_FORM_ID 
                        WHERE i.FIN_SUBJECT_ID in ('6560', '1307') AND f.STATUS = 30 GROUP BY OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID)expense ON budget.SUBJECT_ID + CONVERT(varchar, budget.CURRENT_YEAR) + CONVERT(varchar, budget.BUDGET_MONTH) = expense.FIN_SUBJECT_ID + CONVERT(varchar, expense.OCCURRED_YEAR) + CONVERT(varchar, expense.OCCURRED_MONTH))r GROUP BY r.CURRENT_YEAR,r.BUDGET_MONTH
                        UNION
                        SELECT CONVERT(varchar, r.CURRENT_YEAR) +'/' + CONVERT(varchar, r.BUDGET_MONTH) + '/' + CONVERT(varchar, 5)PaidDate, sum(PaidAmt)AMOUNT from(SELECT budget.*, iif(expense.AMOUNT_REAL is not null, 0, budget.AMOUNT)PaidAmt FROM(SELECT SUBJECT_ID, BUDGET_YEAR AS CURRENT_YEAR, BUDGET_MONTH, AMOUNT FROM PLAN_SITE_BUDGET
                        WHERE SUBJECT_ID IN('1321', '1317', '1318', '1316', '1328', '1315', '1319') AND  BUDGET_YEAR >= YEAR(GETDATE()))budget left join(SELECT OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID, sum(i.AMOUNT) as AMOUNT_REAL FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM i ON f.EXP_FORM_ID = i.EXP_FORM_ID
                        WHERE i.FIN_SUBJECT_ID in ('1321', '1317', '1318', '1316', '1328', '1315', '1319') AND f.STATUS = 30 GROUP BY OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID)expense ON budget.SUBJECT_ID + CONVERT(varchar, budget.CURRENT_YEAR) + CONVERT(varchar, budget.BUDGET_MONTH) = expense.FIN_SUBJECT_ID + CONVERT(varchar, expense.OCCURRED_YEAR) + CONVERT(varchar, expense.OCCURRED_MONTH))r GROUP BY r.CURRENT_YEAR,r.BUDGET_MONTH)exp
                        where exp.AMOUNT <> 0 and CONVERT(char(10), CONVERT(datetime, exp.PaidDate, 111), 111) >= @duringStart and CONVERT(char(10), CONVERT(datetime, exp.PaidDate, 111), 111) <= @duringEnd and CONVERT(char(10), CONVERT(datetime, exp.PaidDate, 111), 111) >= CONVERT(char(10), GETDATE(), 111) ";
                    logger.Debug("sql=" + sql + ",duringStart=" + duringStart + ",duringEnd=" + duringEnd);
                    lstItem = context.Database.SqlQuery<ExpenseFormFunction>(sql, new SqlParameter("duringStart", duringStart), new SqlParameter("duringEnd", duringEnd)).ToList();
                }
                else
                {
                    lstItem = context.Database.SqlQuery<ExpenseFormFunction>("select CONVERT(char(10),CONVERT(datetime, exp.PaidDate , 111), 111)RECORDED_DATE, exp.AMOUNT from (SELECT CONVERT(varchar, r.CURRENT_YEAR) + '/'+ CONVERT(varchar, r.BUDGET_MONTH) + '/'+ CONVERT(varchar, 12)PaidDate, sum(PaidAmt)AMOUNT from (SELECT budget.*, iif(expense.AMOUNT_REAL is not null, 0, budget.AMOUNT)PaidAmt FROM (SELECT SUBJECT_ID, CURRENT_YEAR, BUDGET_MONTH, AMOUNT FROM FIN_EXPENSE_BUDGET WHERE SUBJECT_ID = '6020' AND CURRENT_YEAR >= YEAR(GETDATE()) UNION " +
                    "SELECT SUBJECT_ID, BUDGET_YEAR, BUDGET_MONTH, AMOUNT FROM PLAN_SITE_BUDGET WHERE SUBJECT_ID = '1312' AND BUDGET_YEAR >= YEAR(GETDATE()))budget left join(SELECT OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID, sum(i.AMOUNT) as AMOUNT_REAL FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM i ON f.EXP_FORM_ID = i.EXP_FORM_ID WHERE i.FIN_SUBJECT_ID in ('6020', '1312') AND f.STATUS = 30 GROUP BY OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID)expense " +
                    "ON budget.SUBJECT_ID + CONVERT(varchar, budget.CURRENT_YEAR) + CONVERT(varchar, budget.BUDGET_MONTH) = expense.FIN_SUBJECT_ID + CONVERT(varchar, expense.OCCURRED_YEAR) + CONVERT(varchar, expense.OCCURRED_MONTH))r GROUP BY r.CURRENT_YEAR,r.BUDGET_MONTH " +
                    "UNION SELECT CONVERT(varchar, r.CURRENT_YEAR) +'/' + CONVERT(varchar, r.BUDGET_MONTH) + '/' + CONVERT(varchar, 5)PaidDate, sum(PaidAmt)AMOUNT from(SELECT budget.*, iif(expense.AMOUNT_REAL is not null, 0, budget.AMOUNT)PaidAmt FROM(SELECT SUBJECT_ID, CURRENT_YEAR, BUDGET_MONTH, AMOUNT FROM FIN_EXPENSE_BUDGET WHERE SUBJECT_ID = '6010' AND CURRENT_YEAR >= YEAR(GETDATE()) UNION " +
                    "SELECT SUBJECT_ID, BUDGET_YEAR, BUDGET_MONTH, AMOUNT FROM PLAN_SITE_BUDGET WHERE SUBJECT_ID = '1301' AND  BUDGET_YEAR >= YEAR(GETDATE()))budget left join(SELECT OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID, sum(i.AMOUNT) as AMOUNT_REAL FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM i ON f.EXP_FORM_ID = i.EXP_FORM_ID " +
                    "WHERE i.FIN_SUBJECT_ID in ('6010', '1301') AND f.STATUS = 30 GROUP BY OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID)expense ON budget.SUBJECT_ID + CONVERT(varchar, budget.CURRENT_YEAR) + CONVERT(varchar, budget.BUDGET_MONTH) = expense.FIN_SUBJECT_ID + CONVERT(varchar, expense.OCCURRED_YEAR) + CONVERT(varchar, expense.OCCURRED_MONTH))r GROUP BY r.CURRENT_YEAR,r.BUDGET_MONTH " +
                    "UNION SELECT CONVERT(varchar, r.CURRENT_YEAR) +'/' + CONVERT(varchar, r.BUDGET_MONTH) + '/' + CONVERT(varchar, 15)PaidDate, sum(PaidAmt)AMOUNT from(SELECT budget.*, iif(expense.AMOUNT_REAL is not null, 0, budget.AMOUNT)PaidAmt FROM(SELECT SUBJECT_ID, CURRENT_YEAR, BUDGET_MONTH, AMOUNT FROM FIN_EXPENSE_BUDGET WHERE SUBJECT_ID = '6100' AND CURRENT_YEAR >= YEAR(GETDATE()) UNION " +
                    "SELECT SUBJECT_ID, BUDGET_YEAR, BUDGET_MONTH, AMOUNT FROM PLAN_SITE_BUDGET WHERE SUBJECT_ID = '1306' AND  BUDGET_YEAR >= YEAR(GETDATE()))budget left join(SELECT OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID, sum(i.AMOUNT) as AMOUNT_REAL FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM i ON f.EXP_FORM_ID = i.EXP_FORM_ID " +
                    "WHERE i.FIN_SUBJECT_ID in ('6100', '1306') AND f.STATUS = 30 GROUP BY OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID)expense ON budget.SUBJECT_ID + CONVERT(varchar, budget.CURRENT_YEAR) + CONVERT(varchar, budget.BUDGET_MONTH) = expense.FIN_SUBJECT_ID + CONVERT(varchar, expense.OCCURRED_YEAR) + CONVERT(varchar, expense.OCCURRED_MONTH))r GROUP BY r.CURRENT_YEAR,r.BUDGET_MONTH " +
                    "UNION SELECT CONVERT(varchar, r.CURRENT_YEAR) +'/' + CONVERT(varchar, r.BUDGET_MONTH) + '/' + iif(BUDGET_MONTH <> 2, CONVERT(varchar, 30), CONVERT(varchar, 28))PaidDate, sum(PaidAmt)PaidAmt from(SELECT budget.*, iif(expense.AMOUNT_REAL is not null, 0, budget.AMOUNT)PaidAmt FROM(SELECT SUBJECT_ID, CURRENT_YEAR, BUDGET_MONTH, AMOUNT FROM FIN_EXPENSE_BUDGET WHERE SUBJECT_ID = '6560' AND CURRENT_YEAR >= YEAR(GETDATE()) UNION " +
                    "SELECT SUBJECT_ID, BUDGET_YEAR, BUDGET_MONTH, AMOUNT FROM PLAN_SITE_BUDGET WHERE SUBJECT_ID = '1307' AND  BUDGET_YEAR >= YEAR(GETDATE()))budget left join(SELECT OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID, sum(i.AMOUNT) as AMOUNT_REAL FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM i ON f.EXP_FORM_ID = i.EXP_FORM_ID " +
                    "WHERE i.FIN_SUBJECT_ID in ('6560', '1307') AND f.STATUS = 30 GROUP BY OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID)expense ON budget.SUBJECT_ID + CONVERT(varchar, budget.CURRENT_YEAR) + CONVERT(varchar, budget.BUDGET_MONTH) = expense.FIN_SUBJECT_ID + CONVERT(varchar, expense.OCCURRED_YEAR) + CONVERT(varchar, expense.OCCURRED_MONTH))r GROUP BY r.CURRENT_YEAR,r.BUDGET_MONTH " +
                    "UNION SELECT CONVERT(varchar, r.CURRENT_YEAR) +'/' + CONVERT(varchar, r.BUDGET_MONTH) + '/' + CONVERT(varchar, 5)PaidDate, sum(PaidAmt)AMOUNT from(SELECT budget.*, iif(expense.AMOUNT_REAL is not null, 0, budget.AMOUNT)PaidAmt FROM(SELECT SUBJECT_ID, BUDGET_YEAR AS CURRENT_YEAR, BUDGET_MONTH, AMOUNT FROM PLAN_SITE_BUDGET " +
                    "WHERE SUBJECT_ID IN('1321', '1317', '1318', '1316', '1328', '1315', '1319') AND  BUDGET_YEAR >= YEAR(GETDATE()))budget left join(SELECT OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID, sum(i.AMOUNT) as AMOUNT_REAL FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM i ON f.EXP_FORM_ID = i.EXP_FORM_ID " +
                    "WHERE i.FIN_SUBJECT_ID in ('1321', '1317', '1318', '1316', '1328', '1315', '1319') AND f.STATUS = 30 GROUP BY OCCURRED_YEAR, OCCURRED_MONTH, i.FIN_SUBJECT_ID)expense ON budget.SUBJECT_ID + CONVERT(varchar, budget.CURRENT_YEAR) + CONVERT(varchar, budget.BUDGET_MONTH) = expense.FIN_SUBJECT_ID + CONVERT(varchar, expense.OCCURRED_YEAR) + CONVERT(varchar, expense.OCCURRED_MONTH))r GROUP BY r.CURRENT_YEAR,r.BUDGET_MONTH)exp  " +
                    "where exp.AMOUNT <> 0 and CONVERT(char(10), CONVERT(datetime, exp.PaidDate, 111), 111) >= CONVERT(char(10), GETDATE(), 111) ").ToList();
                }
            }
            return lstItem;
        }

        public int updatePlanAccountItem(PLAN_ACCOUNT item)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.PLAN_ACCOUNT.AddOrUpdate(item);
                    i = context.SaveChanges();
                }
                catch (Exception e)
                {
                    logger.Error("updatePlanAcountItem  fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }
            }
            return i;
        }

        //取得特定年度公司費用每月執行總和
        public List<ExpensetFromOPByMonth> getExpensetOfMonthByYear(int year)
        {
            List<ExpensetFromOPByMonth> lstExpense = new List<ExpensetFromOPByMonth>();
            using (var context = new topmepEntities())
            {
                lstExpense = context.Database.SqlQuery<ExpensetFromOPByMonth>("SELECT SUM(F.JAN) AS JAN, SUM(F.FEB) AS FEB, SUM(F.MAR) AS MAR, SUM(F.APR) AS APR, SUM(F.MAY) AS MAY, SUM(F.JUN) AS JUN, " +
                   "SUM(F.JUL) AS JUL, SUM(F.AUG) AS AUG, SUM(F.SEP) AS SEP, SUM(F.OCT) AS OCT, SUM(F.NOV) AS NOV, SUM(F.DEC) AS DEC, SUM(F.HTOTAL) AS HTOTAL " +
                   "FROM (SELECT C.*, SUM(ISNULL(C.JAN, 0)) + SUM(ISNULL(C.FEB, 0)) + SUM(ISNULL(C.MAR, 0)) + SUM(ISNULL(C.APR, 0)) + SUM(ISNULL(C.MAY, 0)) + SUM(ISNULL(C.JUN, 0)) " +
                    "+ SUM(ISNULL(C.JUL, 0)) + SUM(ISNULL(C.AUG, 0)) + SUM(ISNULL(C.SEP, 0)) + SUM(ISNULL(C.OCT, 0)) + SUM(ISNULL(C.NOV, 0)) + SUM(ISNULL(C.DEC, 0)) AS HTOTAL " +
                    "FROM(SELECT SUBJECT_NAME, FIN_SUBJECT_ID, [01] As 'JAN', [02] As 'FEB', [03] As 'MAR', [04] As 'APR', [05] As 'MAY', [06] As 'JUN', [07] As 'JUL', [08] As 'AUG', [09] As 'SEP', [10] As 'OCT', [11] As 'NOV', [12] As 'DEC' " +
                    "FROM(SELECT B.OCCURRED_MONTH, fs.FIN_SUBJECT_ID, fs.SUBJECT_NAME, B.AMOUNT FROM FIN_SUBJECT fs LEFT JOIN(SELECT ef.OCCURRED_MONTH, ei.FIN_SUBJECT_ID, ei.AMOUNT FROM FIN_EXPENSE_ITEM ei " +
                    "LEFT JOIN FIN_EXPENSE_FORM ef ON ei.EXP_FORM_ID = ef.EXP_FORM_ID WHERE ef.OCCURRED_YEAR = @year AND ef.OCCURRED_MONTH > 6 OR ef.OCCURRED_YEAR = @year + 1 AND ef.OCCURRED_MONTH < 7)B " +
                    "ON fs.FIN_SUBJECT_ID = B.FIN_SUBJECT_ID WHERE fs.CATEGORY = '公司營業費用') As STable " +
                    "PIVOT(SUM(AMOUNT) FOR OCCURRED_MONTH IN([01], [02], [03], [04], [05], [06], [07], [08], [09], [10], [11], [12])) As PTable)C " +
                    "GROUP BY C.SUBJECT_NAME, C.FIN_SUBJECT_ID, C.JAN, C.FEB, C.MAR, C.APR, C.MAY, C.JUN, C.JUL, C.AUG, C.SEP, C.OCT, C.NOV, C.DEC )F ; "
                   , new SqlParameter("year", year)).ToList();
            }
            return lstExpense;
        }
        #endregion

        //取得特定專案之業主計價次數
        public RevenueFromOwner getVACount4OwnerById(string projectid)
        {
            RevenueFromOwner valuation = null;
            using (var context = new topmepEntities())
            {
                valuation = context.Database.SqlQuery<RevenueFromOwner>("SELECT DISTINCT p.PROJECT_ID, ISNULL((SELECT COUNT(PROJECT_ID) FROM PLAN_VALUATION_FORM " +
                    "WHERE VALUATION_AMOUNT IS NOT NULL GROUP BY PROJECT_ID HAVING PROJECT_ID =@projectid),0)+1 AS VACount, ISNULL((SELECT COUNT(PROJECT_ID) FROM PLAN_VALUATION_FORM " +
                    "GROUP BY PROJECT_ID HAVING PROJECT_ID =@projectid),0)+1 AS isVA FROM TND_PROJECT p WHERE p.PROJECT_ID =@projectid "
                   , new SqlParameter("projectid", projectid)).First();
            }
            return valuation;
        }

        public string refreshVA(string formid, PLAN_VALUATION_FORM form)
        {
            int i = 0;
            if (null != formid && formid != "")
            {
                form.VA_FORM_ID = formid;
            }
            else
            {
                string sno_key = "VA";
                SerialKeyService snoservice = new SerialKeyService();
                form.VA_FORM_ID = snoservice.getSerialKey(sno_key);
            }
            using (var context = new topmepEntities())
            {
                try
                {
                    context.PLAN_VALUATION_FORM.AddOrUpdate(form);
                    i = context.SaveChanges();
                }
                catch (Exception e)
                {
                    logger.Error("update VA item fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return form.VA_FORM_ID;
        }

        /*
      public RevenueFromOwner getVAPayItemById(string formid)
      {
          RevenueFromOwner payment = null;
          using (var context = new topmepEntities())
          {
              payment = context.Database.SqlQuery<RevenueFromOwner>("SELECT A.VA_FORM_ID , ROUND(CAST(IIF(ISNULL(A.advancePaymentBalance, 0) - ISNULL(A.VALUATION_AMOUNT , 0)*A.ADVANCE_RATIO/100 > 0, " +
                  "ISNULL(A.VALUATION_AMOUNT , 0)*A.ADVANCE_RATIO/100, IIF(ISNULL(A.advancePaymentBalance, 0) > 0, A.advancePaymentBalance, 0)) AS decimal(10,1)),0) AS ADVANCE_PAYMENT_REFUND, " +
                  "ROUND(CAST(ISNULL(A.VALUATION_AMOUNT , 0)*A.RETENTION_RATIO/100 AS decimal(10,1)),0) AS RETENTION_PAYMENT, ROUND(CAST((ISNULL(A.VALUATION_AMOUNT , 0)-ISNULL(A.VALUATION_AMOUNT , 0) " +
                  "*A.ADVANCE_RATIO/100)*ISNULL(A.TAX_RATIO,0) / 100 AS decimal(10,1)),0) AS TAX_AMOUNT " +
                  "FROM (SELECT vf.*, B.advancePaymentBalance, ISNULL(IIF(ppt.PAYMENT_TERMS = 'P', ppt.PAYMENT_ADVANCE_RATIO, ppt.USANCE_ADVANCE_RATIO), 0) AS ADVANCE_RATIO, " +
                  "ISNULL(IIF(ppt.PAYMENT_TERMS = 'P', ppt.PAYMENT_RETENTION_RATIO, ppt.USANCE_RETENTION_RATIO),0) AS RETENTION_RATIO FROM PLAN_VALUATION_FORM vf " +
                  "LEFT JOIN PLAN_PAYMENT_TERMS ppt ON vf.PROJECT_ID = ppt.CONTRACT_ID LEFT JOIN (SELECT PROJECT_ID, SUM(ADVANCE_PAYMENT) - SUM(ADVANCE_PAYMENT_REFUND) AS advancePaymentBalance " +
                  "FROM PLAN_VALUATION_FORM GROUP BY PROJECT_ID)B ON vf.PROJECT_ID = B.PROJECT_ID WHERE vf.VA_FORM_ID =@formid)A  "
                 , new SqlParameter("formid", formid)).FirstOrDefault();
          }
          return payment;
      }

      //寫入保留款,預付扣回與營業稅額(舊)
      public int refreshVAItem(string formid, decimal retention, decimal advanceRefund, decimal tax)
      {
          int i = 0;
          logger.Info("refresh rentention amount, advance refund and tax amount of VA by form id" + formid);
          string sql = "UPDATE  PLAN_VALUATION_FORM SET RETENTION_PAYMENT =@retention, ADVANCE_PAYMENT_REFUND =@advanceRefund, TAX_AMOUNT =@tax WHERE VA_FORM_ID=@formid ";
          logger.Debug("refresh items from Quote sql:" + sql);
          db = new topmepEntities();
          var parameters = new List<SqlParameter>();
          parameters.Add(new SqlParameter("formid", formid));
          parameters.Add(new SqlParameter("retention", retention));
          parameters.Add(new SqlParameter("advanceRefund", advanceRefund));
          parameters.Add(new SqlParameter("tax", tax));
          db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
          i = db.SaveChanges();
          logger.Info("Update Record:" + i);
          db = null;
          return i;
      }
      */
        public List<RevenueFromOwner> getVADetailById(string projectid)
        {
            List<RevenueFromOwner> VAItem = new List<RevenueFromOwner>();
            using (var context = new topmepEntities())
            {
                VAItem = context.Database.SqlQuery<RevenueFromOwner>("SELECT vf.*, pi.discount, (pi.taxAmt - pi.taxMinus) AS taxAmt , ISNULL(account.AR_PAID, 0) AS AR_PAID, f.FILE_UPLOAD_NAME, ISNULL(vf.ADVANCE_PAYMENT, 0) + ISNULL(vf.VALUATION_AMOUNT,0) + pi.taxAmt " +
                    "- ISNULL(vf.RETENTION_PAYMENT, 0) - ISNULL(vf.ADVANCE_PAYMENT_REFUND, 0) - pi.discount - pi.otherPay - pi.taxMinus AS AR, " +
                    "ISNULL(vf.ADVANCE_PAYMENT, 0) + ISNULL(vf.VALUATION_AMOUNT,0) + pi.taxAmt - ISNULL(vf.RETENTION_PAYMENT, 0) - ISNULL(vf.ADVANCE_PAYMENT_REFUND, 0) - pi.discount " +
                    "- pi.otherPay - ISNULL(account.AR_PAID, 0) - pi.taxMinus AS AR_UNPAID, " +
                    "ROW_NUMBER() OVER(ORDER BY vf.CREATE_DATE) AS NO FROM PLAN_VALUATION_FORM vf LEFT JOIN (SELECT pa.ACCOUNT_FORM_ID, SUM(pa.AMOUNT_PAYABLE) AS AR_PAID FROM PLAN_ACCOUNT pa " +
                    "WHERE pa.ACCOUNT_TYPE = 'R' AND pa.PROJECT_ID =@projectid AND pa.STATUS = 10 GROUP BY pa.ACCOUNT_FORM_ID)account ON vf.VA_FORM_ID = account.ACCOUNT_FORM_ID " +
                    "LEFT JOIN (SELECT FILE_UPLOAD_NAME FROM TND_FILE GROUP BY FILE_UPLOAD_NAME)f ON vf.VA_FORM_ID = f.FILE_UPLOAD_NAME " +
                    "LEFT JOIN (SELECT EST_FORM_ID, ISNULL(SUM(AMOUNT*IIF(TYPE <> '折讓單', 0, 1)), 0) + ISNULL(SUM(TAX * IIF(TYPE <> '折讓單', 0, 1)), 0) AS discount, " +
                    "ISNULL(SUM(TAX * IIF(TYPE = '其他扣款', 1, 0)), 0) AS taxMinus, " +
                    "ISNULL(SUM(AMOUNT * IIF(TYPE = '其他扣款', 1, 0)), 0) AS otherPay, " +
                    "ISNULL(SUM(TAX*IIF(TYPE <> '折讓單', 1, 0)), 0) AS taxAmt FROM PLAN_INVOICE GROUP BY EST_FORM_ID)pi " +
                    "ON pi.EST_FORM_ID = vf.VA_FORM_ID WHERE vf.PROJECT_ID =@projectid "
                   , new SqlParameter("projectid", projectid)).ToList();
            }
            return VAItem;
        }
        public RevenueFromOwner getVADetailByVAId(string formid)
        {
            RevenueFromOwner detail = null;
            using (var context = new topmepEntities())
            {
                detail = context.Database.SqlQuery<RevenueFromOwner>("SELECT vf.*, CONVERT(varchar, vf.CREATE_DATE, 120) AS RECORDED_DATE, " +
                    "ISNULL(vf.ADVANCE_PAYMENT, 0) + ISNULL(vf.VALUATION_AMOUNT, 0) + pi.taxAmt - ISNULL(vf.RETENTION_PAYMENT, 0) - ISNULL(vf.ADVANCE_PAYMENT_REFUND, 0) - pi.otherPay AS AR " +
                    "FROM PLAN_VALUATION_FORM vf LEFT JOIN (SELECT EST_FORM_ID, ISNULL(SUM(AMOUNT * IIF(TYPE <> '折讓單', 0, 1)), 0) + ISNULL(SUM(TAX * IIF(TYPE <> '折讓單', 0, 1)), 0) AS otherPay, ISNULL(SUM(TAX * IIF(TYPE <> '折讓單', 1, 0)), 0) AS taxAmt " +
                    "FROM PLAN_INVOICE GROUP BY EST_FORM_ID)pi ON pi.EST_FORM_ID = vf.VA_FORM_ID WHERE vf.VA_FORM_ID =@formid  "
                   , new SqlParameter("formid", formid)).FirstOrDefault();
            }
            return detail;
        }

        public RevenueFromOwner getVASummaryAtmById(string prjid)
        {
            RevenueFromOwner summaryAmt = null;
            string sql = @"SELECT vf.*, pi.otherPay, pi.taxAmt, vf.Amt + pi.taxAmt - pi.otherPay AS AR , 
                   (SELECT SUM(ITEM_UNIT_PRICE*ITEM_QUANTITY) FROM PLAN_ITEM pi WHERE pi.PROJECT_ID =@pid) AS contractAtm, 
                   (SELECT SUM(pa.AMOUNT_PAYABLE) FROM PLAN_ACCOUNT pa WHERE pa.ACCOUNT_TYPE = 'R' 
                    AND pa.PROJECT_ID =@pid AND pa.STATUS = 10) AS AR_PAID FROM (
                    SELECT PROJECT_ID, SUM(VALUATION_AMOUNT) AS VALUATION_AMOUNT, SUM(RETENTION_PAYMENT) AS RETENTION_PAYMENT, 
                   isnull(SUM(ADVANCE_PAYMENT), 0) - isnull(SUM(ADVANCE_PAYMENT_REFUND), 0) AS advancePaymentBalance, 
                   isnull(SUM(ADVANCE_PAYMENT), 0) + isnull(SUM(VALUATION_AMOUNT), 0) - isnull(SUM(RETENTION_PAYMENT), 0) - isnull(SUM(ADVANCE_PAYMENT_REFUND), 0) AS Amt 
                   FROM PLAN_VALUATION_FORM WHERE PROJECT_ID =@pid GROUP BY PROJECT_ID
                   )vf 
                    LEFT JOIN(
                    SELECT CONTRACT_ID, ISNULL(SUM(AMOUNT * IIF(TYPE <> '折讓單', 0, 1)), 0) 
                    + ISNULL(SUM(TAX * IIF(TYPE <> '折讓單', 0, 1)), 0) 
                    + ISNULL(SUM(AMOUNT * IIF(TYPE <> '其他扣款', 0, 1)), 0) AS otherPay, 
                   ISNULL(SUM(TAX * IIF(TYPE <> '折讓單', 1, 0)), 0) - ISNULL(SUM(TAX * IIF(TYPE <> '其他扣款', 0, 1)), 0) AS taxAmt 
                   FROM PLAN_INVOICE WHERE CONTRACT_ID =@pid GROUP BY CONTRACT_ID)pi ON vf.PROJECT_ID = pi.CONTRACT_ID ";
            using (var context = new topmepEntities())
            {
                summaryAmt = context.Database.SqlQuery<RevenueFromOwner>(sql , new SqlParameter("pid", prjid)).FirstOrDefault();
            }
            return summaryAmt;
        }
        //寫入應收帳款支付資料
        public int addPlanAccount(PLAN_ACCOUNT form)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.PLAN_ACCOUNT.AddOrUpdate(form);
                    i = context.SaveChanges();
                }
                catch (Exception e)
                {
                    logger.Error("update plan account item fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }
        //取得特定專案銀行貸款資料
        public void getAllBankLoan(string projectid)
        {
            List<FIN_BANK_LOAN> lstLoans = null;
            using (var context = new topmepEntities())
            {
                try
                {
                    lstLoans = context.FIN_BANK_LOAN.SqlQuery("SELECT bl.*, it.loanAmt AS LOAN FROM FIN_BANK_LOAN bl LEFT JOIN " +
                        "(SELECT flt.BL_ID, ISNULL(SUM(flt.AMOUNT*flt.TRANSACTION_TYPE), 0) AS loanAmt FROM FIN_LOAN_TRANACTION flt GROUP BY flt.BL_ID)it " +
                        "ON bl.BL_ID = it.BL_ID WHERE bl.PROJECT_ID = @projectid AND it.loanAmt <> 0 ", new SqlParameter("projectid", projectid)).ToList();
                    logger.Debug("get records=" + lstLoans.Count);
                    //將特定專案之貸款封裝供前端頁面調用
                    cashFlowModel.finLoan = lstLoans;
                }
                catch (Exception e)
                {
                    logger.Error("fail:" + e.StackTrace);
                }
            }
        }
        public FIN_BANK_LOAN getPaybackRatioById(int blid)
        {
            FIN_BANK_LOAN loan = null;
            using (var context = new topmepEntities())
            {
                loan = context.FIN_BANK_LOAN.SqlQuery("select loan.* from FIN_BANK_LOAN loan "
                    + "where loan.BL_ID = @blid "
                   , new SqlParameter("blid", blid)).FirstOrDefault();
            }
            return loan;
        }
        //取得貸款還款次數
        public int getPaybackCountByBlId(int blid)
        {
            int count = 0;
            using (var context = new topmepEntities())
            {
                count = context.Database.SqlQuery<int>("SELECT ISNULL((SELECT COUNT(*) FROM FIN_LOAN_TRANACTION " +
                    "WHERE BL_ID = @blid),0)+1 AS paybackCount  "
                   , new SqlParameter("blid", blid)).First();
            }
            return count;
        }
        //取得借款還款餘額
        public decimal getLoanBalanceByBlId(int blid)
        {
            decimal balance = 0;
            using (var context = new topmepEntities())
            {
                balance = context.Database.SqlQuery<decimal>("SELECT ISNULL(SUM(AMOUNT * flt.TRANSACTION_TYPE), 0) AS loanBalance FROM FIN_LOAN_TRANACTION flt WHERE BL_ID =@blid  "
                   , new SqlParameter("blid", blid)).First();
            }
            return balance;
        }
        //更新償還貸款金額以入各專案之備償專戶
        public int addLoanTransaction(int loanid, decimal atm, DateTime paybackDate, string createId, int period, string loanRemark, string formid)
        {
            int i = 0;
            logger.Info("pay back to loan account by loan id" + loanid);
            string sql = "INSERT INTO FIN_LOAN_TRANACTION (BL_ID, TRANSACTION_TYPE, AMOUNT, PAYBACK_DATE, CREATE_ID, CREATE_DATE, PERIOD, REMARK, VA_FORM_ID) VALUES (@loanid, 1, @atm, @paybackDate,@createId, getdate(),@period, @loanRemark, @formid) ";
            logger.Debug("batch sql:" + sql);
            db = new topmepEntities();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("loanid", loanid));
            parameters.Add(new SqlParameter("atm", atm));
            parameters.Add(new SqlParameter("paybackDate", paybackDate));
            parameters.Add(new SqlParameter("createId", createId));
            parameters.Add(new SqlParameter("period", period));
            parameters.Add(new SqlParameter("loanRemark", loanRemark));
            parameters.Add(new SqlParameter("formid", formid));
            db.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            i = db.SaveChanges();
            logger.Info("Update Record:" + i);
            db = null;
            return 1;
        }
        //取得上傳的計價檔案紀錄
        public List<RevenueFromOwner> getVAFileByFormId(string formid)
        {

            logger.Info(" get VA file by form id =" + formid);
            List<RevenueFromOwner> lstItem = new List<RevenueFromOwner>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT f.FILE_ID AS ITEM_UID, f.PROJECT_ID, f.FILE_UPLOAD_NAME, f.FILE_ACTURE_NAME, f.FILE_TYPE, " +
                "f.CREATE_DATE, ROW_NUMBER() OVER(ORDER BY FILE_ID) AS NO FROM TND_FILE f WHERE f.FILE_UPLOAD_NAME = @formid ";

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formid", formid));
            using (var context = new topmepEntities())
            {
                logger.Debug("get VA file sql=" + sql);
                lstItem = context.Database.SqlQuery<RevenueFromOwner>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get VA file record count=" + lstItem.Count);
            return lstItem;
        }

        //取得各專案財務執行進度
        public List<PlanFinanceProfile> getPlanFinProfile()
        {

            logger.Info("get the finance profile of all plans !! ");
            List<PlanFinanceProfile> lstItem = new List<PlanFinanceProfile>();
            string sql = @"
WITH IDX_PRJ AS (
--執行中專案
SELECT * FROM TND_PROJECT WHERE STATUS='專案執行')
,CONTRACT_AMOUNT AS (
--簽約金額-對業主
SELECT PROJECT_ID,SUM(ISNULL(ITEM_QUANTITY,0) * ISNULL(ITEM_UNIT_PRICE,0)) CONTRACT_AMOUNT
FROM PLAN_ITEM WHERE PROJECT_ID 
IN (SELECT PROJECT_ID FROM IDX_PRJ)
GROUP BY PROJECT_ID 
),
CONTRACT_TENDER AS (
--專案成本-發包
SELECT F.PROJECT_ID,SUM(ISNULL(I.ITEM_QTY,0) * ISNULL(I.ITEM_UNIT_PRICE,0)) TENDER_AMOUNT
 FROM PLAN_SUP_INQUIRY F,PLAN_SUP_INQUIRY_ITEM I
WHERE F.INQUIRY_FORM_ID=I.INQUIRY_FORM_ID
AND F.PROJECT_ID IN 
(SELECT PROJECT_ID FROM IDX_PRJ)
GROUP BY F.PROJECT_ID
),
SITE_BUDGET AS (
--專案成本-工地預算
SELECT PROJECT_ID,SUM(AMOUNT) SITE_BUDGET_AMOUNT FROM PLAN_SITE_BUDGET
WHERE PROJECT_ID IN 
(SELECT PROJECT_ID FROM IDX_PRJ)
GROUP BY PROJECT_ID
),
OTHER_COST AS (
--專案成本-其他成本
SELECT PROJECT_ID,SUM(COST) OTHER_COST 
FROM PLAN_INDIRECT_COST 
WHERE PROJECT_ID IN  (SELECT PROJECT_ID FROM IDX_PRJ)
AND FIELD_ID IN ('04_SiteCost','02_SalesCost','01_MACost')
GROUP BY PROJECT_ID
),
AP AS (
SELECT 
PROJECT_ID,
SUM(ISNULL(ADVANCE_PAYMENT,0) + ISNULL(VALUATION_AMOUNT,0) 
- ISNULL(ADVANCE_PAYMENT_REFUND,0)
-ISNULL(RETENTION_PAYMENT,0) + ISNULL(TAX_AMOUNT,0)) AP_AMOUNT,
(SELECT SUM(ISNULL(AMOUNT,0)+ISNULL(TAX,0))
     DEDUCTION_AMOUNT  FROM PLAN_INVOICE I WHERE TYPE IN('折讓單','其他扣款')
	 AND F.PROJECT_ID=I.CONTRACT_ID GROUP BY CONTRACT_ID 
	) DEDUCTION_AMOUNT
 FROM PLAN_VALUATION_FORM F
 WHERE F.VA_FORM_ID IN 
 (SELECT ACCOUNT_FORM_ID FROM PLAN_ACCOUNT)
 GROUP BY PROJECT_ID
)
SELECT P.PROJECT_ID,P.PROJECT_NAME, 
C.CONTRACT_AMOUNT * 1.05 AS CONTRACT_AMOUNT,
T.TENDER_AMOUNT * 1.05 AS TENDER_AMOUNT
,B.SITE_BUDGET_AMOUNT,O.OTHER_COST,
(ISNULL(AP.AP_AMOUNT,0)-ISNULL(AP.DEDUCTION_AMOUNT,0)) * 1.05 AS  AR_AMOUNT
FROM IDX_PRJ P 
LEFT JOIN CONTRACT_AMOUNT C ON P.PROJECT_ID=C.PROJECT_ID
LEFT OUTER JOIN CONTRACT_TENDER T ON P.PROJECT_ID=T.PROJECT_ID
LEFT OUTER JOIN SITE_BUDGET B ON P.PROJECT_ID=B.PROJECT_ID
LEFT OUTER JOIN OTHER_COST O ON P.PROJECT_ID=O.PROJECT_ID
LEFT OUTER JOIN AP ON P.PROJECT_ID=AP.PROJECT_ID
;";
            //處理SQL 預先填入ID,設定集合處理參數
            using (var context = new topmepEntities())
            {
                lstItem = context.Database.SqlQuery<PlanFinanceProfile>(sql).ToList();
                logger.Info("Get Plan Financial Profile Count=" + lstItem.Count);
            }

            return lstItem;
        }
        //取得所有專案財務執行進度
        public PlanFinanceProfile getFinProfile()
        {
            string sql = "SELECT SUM(PLAN_REVENUE) AS PLAN_REVENUE, SUM(directCost) AS directCost, " +
                    "SUM(AP) AS AP, SUM(AR) AS AR, SUM(uncollectedAR) AS uncollectedAR, SUM(unpaidAP) AS unpaidAP, SUM(MACost) AS MACost, " +
                    "SUM(SiteCost) AS SiteCost, SUM(SalesCost) AS SalesCost, SUM(SiteCostPaid) AS SiteCostPaid, SUM(MACostPaid) AS MACostPaid, SUM(OtherCostPaid) AS OtherCostPaid, " +
                    "SUM(ManagementCost) AS ManagementCost, SUM(planProfit) AS planProfit FROM (SELECT *, uncollectedAR - unpaidAP - IIF(SiteCost - SiteCostPaid >= 0, SiteCost - SiteCostPaid, SiteCostPaid - SiteCost) " +
                    "- ManagementCost - IIF(MACost - MACostPaid >= 0, MACost - MACostPaid, MACostPaid - MACost) - SalesCost - OtherCostPaid AS planProfit " +
                    "FROM (SELECT p.PROJECT_ID, p.PROJECT_NAME, ISNULL(B.directCost, 0) AS directCost, ISNULL(C.AR, 0) AS AR, " +
                    "ISNULL(SUM(pi.ITEM_UNIT_PRICE * pi.ITEM_QUANTITY), 0) - ISNULL(C.AR, 0) AS uncollectedAR, ISNULL(D.AP, 0) - ISNULL(E.MinusItem, 0) AS AP, " +
                    "ISNULL(B.directCost, 0) - ISNULL(D.AP, 0) + ISNULL(E.MinusItem, 0) AS unpaidAP, ISNULL(SUM(pi.ITEM_UNIT_PRICE * pi.ITEM_QUANTITY), 0) AS PLAN_REVENUE, " +
                    "ISNULL(F.MACost, 0) AS MACost, ISNULL(G.SalesCost, 0) AS SalesCost, ISNULL(I.CompanyCost, 0) AS ManagementCost, ISNULL(H.SiteCost, 0) AS SiteCost, " +
                    "ISNULL(J.SiteCostPaid, 0) AS SiteCostPaid, ISNULL(K.MACostPaid, 0) AS MACostPaid, ISNULL(M.OtherCostPaid, 0) AS OtherCostPaid " +
                    "FROM TND_PROJECT p LEFT JOIN PLAN_ITEM pi ON pi.PROJECT_ID = p.PROJECT_ID LEFT JOIN (SELECT main.PROJECT_ID, SUM(sub.directCost) AS directCost " +
                    "FROM(SELECT pi.INQUIRY_FORM_ID, pi.PROJECT_ID FROM PLAN_ITEM pi union SELECT pi.MAN_FORM_ID, pi.PROJECT_ID FROM PLAN_ITEM pi)main " +
                    "LEFT JOIN(SELECT psi.INQUIRY_FORM_ID, ROUND(SUM(psi.ITEM_UNIT_PRICE * psi.ITEM_QTY), 0) AS directCost FROM  PLAN_SUP_INQUIRY_ITEM psi GROUP BY psi.INQUIRY_FORM_ID)sub " +
                    "ON main.INQUIRY_FORM_ID = sub.INQUIRY_FORM_ID GROUP BY main.PROJECT_ID)B ON p.PROJECT_ID = B.PROJECT_ID " +
                    "LEFT JOIN (SELECT PROJECT_ID, ISNULL(SUM(ADVANCE_PAYMENT), 0) + ISNULL(SUM(VALUATION_AMOUNT), 0) - ISNULL(SUM(RETENTION_PAYMENT), 0) - ISNULL(SUM(ADVANCE_PAYMENT_REFUND), 0) AS AR " +
                    "FROM PLAN_VALUATION_FORM GROUP BY PROJECT_ID UNION SELECT (SELECT PROJECT_ID FROM TND_PROJECT WHERE PROJECT_ID = '001')PROJECT_ID, SUM(AMOUNT) AS CompanyCostPaid " +
                    "FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM fi ON f.EXP_FORM_ID = fi.EXP_FORM_ID WHERE ISNULL(PROJECT_ID, '') = '' AND STATUS = 30 AND OCCURRED_YEAR = IIF(MONTH(GETDATE()) > 6, YEAR(GETDATE()), YEAR(GETDATE())-1) " +
                    "AND OCCURRED_MONTH > 6 OR ISNULL(PROJECT_ID, '') = '' AND STATUS = 30 AND OCCURRED_YEAR = IIF(MONTH(GETDATE()) > 6, YEAR(GETDATE())+1, YEAR(GETDATE())) AND OCCURRED_MONTH < 6 GROUP BY PROJECT_ID)C " +
                    "ON p.PROJECT_ID = C.PROJECT_ID LEFT JOIN (SELECT PROJECT_ID, ISNULL(SUM(PAYMENT_TRANSFER), 0) - ISNULL(SUM(RETENTION_PAYMENT), 0) AS AP " +
                    "FROM PLAN_ESTIMATION_FORM WHERE ISNULL(INDIRECT_COST_TYPE, '') = '' GROUP BY PROJECT_ID)D ON p.PROJECT_ID = D.PROJECT_ID LEFT JOIN(SELECT PROJECT_ID, ISNULL(SUM(pop.AMOUNT),0) AS MinusItem FROM PLAN_ESTIMATION_FORM pef " +
                    "LEFT JOIN PLAN_OTHER_PAYMENT pop ON pef.EST_FORM_ID = pop.EST_FORM_ID WHERE pop.TYPE  in ('A', 'B', 'C', 'F') GROUP BY PROJECT_ID)E ON p.PROJECT_ID = E.PROJECT_ID " +
                    "LEFT JOIN (SELECT PROJECT_ID, COST AS MACost FROM PLAN_INDIRECT_COST WHERE FIELD_ID = '01_MACost')F on p.PROJECT_ID = F.PROJECT_ID " +
                    "LEFT JOIN (SELECT PROJECT_ID, COST AS SalesCost FROM PLAN_INDIRECT_COST WHERE FIELD_ID = '02_SalesCost')G on p.PROJECT_ID = G.PROJECT_ID " +
                    "LEFT JOIN (SELECT PROJECT_ID, SUM(AMOUNT) AS SiteCost FROM PLAN_SITE_BUDGET GROUP BY PROJECT_ID)H on p.PROJECT_ID = H.PROJECT_ID " +
                    "LEFT JOIN (SELECT PROJECT_ID, SUM(AMOUNT) AS SiteCostPaid FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM fi ON f.EXP_FORM_ID = fi.EXP_FORM_ID WHERE PROJECT_ID IS NOT NULL AND PROJECT_ID <> '' AND STATUS = 30 GROUP BY PROJECT_ID)J on p.PROJECT_ID = J.PROJECT_ID " +
                    "LEFT JOIN (SELECT PROJECT_ID, ISNULL(SUM(PAYMENT_TRANSFER),0) - ISNULL(SUM(RETENTION_PAYMENT),0) AS MACostPaid " +
                    "FROM PLAN_ESTIMATION_FORM WHERE ISNULL(INDIRECT_COST_TYPE, '') = 'M' GROUP BY PROJECT_ID)K ON p.PROJECT_ID = K.PROJECT_ID " +
                    "LEFT JOIN (SELECT PROJECT_ID, ISNULL(SUM(PAYMENT_TRANSFER),0) - ISNULL(SUM(RETENTION_PAYMENT),0) AS OtherCostPaid " +
                    "FROM PLAN_ESTIMATION_FORM WHERE ISNULL(INDIRECT_COST_TYPE, '') = 'O' GROUP BY PROJECT_ID)M ON p.PROJECT_ID = M.PROJECT_ID " +
                    "LEFT JOIN (SELECT PROJECT_ID, COST AS CompanyCost FROM PLAN_INDIRECT_COST WHERE FIELD_ID = '03_CompanyCost')I on p.PROJECT_ID = I.PROJECT_ID WHERE p.STATUS = '專案執行' " +
                    "GROUP BY p.PROJECT_ID, p.PROJECT_NAME, B.directCost, C.AR, D.AP, E.MinusItem, F.MACost, G.SalesCost, I.CompanyCost, H.SiteCost, J.SiteCostPaid,K.MACostPaid, M.OtherCostPaid)a)it ";
            PlanFinanceProfile lstAmount = null;
            using (var context = new topmepEntities())
            {
                lstAmount = context.Database.SqlQuery<PlanFinanceProfile>(sql).FirstOrDefault();
            }
            return lstAmount;
        }
        //取得公司當前現金餘額
        public CashFlowBalance getCashFlowBalance()
        {
            CashFlowBalance lstAmount = null;
            using (var context = new topmepEntities())
            {
                lstAmount = context.Database.SqlQuery<CashFlowBalance>("SELECT *, ISNULL(curCashFlow,0) - ISNULL(loanBalance_sup, 0) + ISNULL(loanBalance_bank, 0) + ISNULL(futureCashFlow, 0) - ISNULL(CompanyCost, 0) AS cashFlowBal FROM " +
                    "(SELECT (SELECT ISNULL(bankAmt, 0) + ISNULL(cashFlow, 0) + ISNULL(loanFlowFactor, 0) - ISNULL(expenseAmt, 0) FROM (SELECT (SELECT SUM(CUR_AMOUNT) FROM FIN_BANK_ACCOUNT) AS bankAmt, " +
                    "(SELECT SUM(fei.AMOUNT) FROM FIN_EXPENSE_FORM fef LEFT JOIN FIN_EXPENSE_ITEM fei ON fef.EXP_FORM_ID = fei.EXP_FORM_ID WHERE fef.STATUS = 30 AND CONVERT(char(10), fef.PAYMENT_DATE, 111)  >= CONVERT(char(10), GETDATE(), 111)) AS expenseAmt, " +
                    "(SELECT SUM(IIF(ISNULL(IS_SUPPLIER, 'N') = 'Y', IIF(TRANSACTION_TYPE = 1, 1 ,-1),IIF(TRANSACTION_TYPE = 1, -1 ,1)) * AMOUNT) " +
                    "FROM FIN_LOAN_TRANACTION t LEFT JOIN FIN_BANK_LOAN l ON t.BL_ID = l.BL_ID WHERE CONVERT(char(10), IIF(TRANSACTION_TYPE = 1, PAYBACK_DATE, EVENT_DATE),111) >= " +
                    "CONVERT(char(10), getdate(),111) AND t.REMARK NOT LIKE '%備償款%') AS loanFlowFactor, SUM(AMOUNT_PAID *IIF(pa.ISDEBIT = 'N', -1, 1)) AS cashFlow " +
                    "FROM PLAN_ACCOUNT pa WHERE pa.PAYMENT_DATE >= CONVERT(char(10), getdate(), 111) AND ISNULL(STATUS, 10) <> 0)it) AS curCashFlow, (SELECT SUM(MAINTENANCE_BOND) FROM PLAN_CONTRACT_PROCESS) AS maintBond, " +
                    "(SELECT SUM(AMOUNT) FROM FIN_EXPENSE_BUDGET WHERE BUDGET_YEAR = IIF(MONTH(GETDATE()) > 6, YEAR(GETDATE()), YEAR(GETDATE())-1)) AS CompanyCost," +
                    "(SELECT ISNULL(SUM(AMOUNT * flt.TRANSACTION_TYPE), 0) FROM FIN_LOAN_TRANACTION flt LEFT JOIN FIN_BANK_LOAN fbl ON flt.BL_ID = fbl.BL_ID WHERE ISNULL(IS_SUPPLIER, 'N') <> 'Y') AS loanBalance_bank, " +
                    "(SELECT ISNULL(SUM(AMOUNT * flt.TRANSACTION_TYPE), 0) FROM FIN_LOAN_TRANACTION flt LEFT JOIN FIN_BANK_LOAN fbl ON flt.BL_ID = fbl.BL_ID WHERE ISNULL(IS_SUPPLIER, 'N') = 'Y') AS loanBalance_sup, " +
                    "(SELECT SUM(ISNULL(uncollectedAR, 0)) - SUM(unpaidAP) - SUM(ISNULL(IIF(SiteCost - SiteCostPaid >= 0,SiteCost - SiteCostPaid, SiteCostPaid - SiteCost), 0)) - " +
                    "SUM(ISNULL(ManagementCost, 0)) - SUM(ISNULL(IIF(MACost - MACostPaid >= 0, MACost - MACostPaid, MACostPaid - MACost), 0)) - ISNULL(SUM(OtherCostPaid), 0) " +
                    "FROM(SELECT p.PROJECT_ID, p.PROJECT_NAME, ISNULL(B.directCost, 0) AS directCost, ISNULL(C.AR, 0) AS AR, " +
                    "ISNULL(SUM(pi.ITEM_UNIT_PRICE * pi.ITEM_QUANTITY), 0) - ISNULL(C.AR, 0) AS uncollectedAR, ISNULL(D.AP, 0) - ISNULL(E.MinusItem, 0) AS AP, " +
                    "ISNULL(B.directCost, 0) - ISNULL(D.AP, 0) + ISNULL(E.MinusItem, 0) AS unpaidAP, ISNULL(SUM(pi.ITEM_UNIT_PRICE * pi.ITEM_QUANTITY), 0) AS PLAN_REVENUE, " +
                    "ISNULL(F.MACost, 0) AS MACost, ISNULL(G.SalesCost, 0) + ISNULL(I.CompanyCost, 0) AS ManagementCost, ISNULL(H.SiteCost, 0) AS SiteCost, " +
                    "ISNULL(J.SiteCostPaid, 0) AS SiteCostPaid, ISNULL(K.MACostPaid, 0) AS MACostPaid, ISNULL(M.OtherCostPaid, 0) AS OtherCostPaid " +
                    "FROM PLAN_ITEM pi LEFT JOIN TND_PROJECT p ON pi.PROJECT_ID = p.PROJECT_ID LEFT JOIN(SELECT main.PROJECT_ID, SUM(sub.directCost) AS directCost " +
                    "FROM (SELECT pi.INQUIRY_FORM_ID, pi.PROJECT_ID FROM PLAN_ITEM pi union SELECT pi.MAN_FORM_ID, pi.PROJECT_ID FROM PLAN_ITEM pi)main " +
                    "LEFT JOIN (SELECT psi.INQUIRY_FORM_ID, ROUND(SUM(psi.ITEM_UNIT_PRICE * psi.ITEM_QTY), 0) AS directCost FROM  PLAN_SUP_INQUIRY_ITEM psi GROUP BY psi.INQUIRY_FORM_ID)sub " +
                    "ON main.INQUIRY_FORM_ID = sub.INQUIRY_FORM_ID GROUP BY main.PROJECT_ID)B ON p.PROJECT_ID = B.PROJECT_ID " +
                    "LEFT JOIN (SELECT PROJECT_ID, ISNULL(SUM(ADVANCE_PAYMENT), 0) + ISNULL(SUM(VALUATION_AMOUNT), 0) - ISNULL(SUM(RETENTION_PAYMENT), 0) - ISNULL(SUM(ADVANCE_PAYMENT_REFUND), 0) AS AR " +
                    "FROM PLAN_VALUATION_FORM GROUP BY PROJECT_ID UNION SELECT (SELECT PROJECT_ID FROM TND_PROJECT WHERE PROJECT_NAME LIKE '%公司費用%')PROJECT_ID, SUM(AMOUNT) AS CompanyCostPaid " +
                    "FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM fi ON f.EXP_FORM_ID = fi.EXP_FORM_ID WHERE ISNULL(PROJECT_ID, '') = '' AND STATUS = 30 AND OCCURRED_YEAR = IIF(MONTH(GETDATE()) > 6, YEAR(GETDATE()), YEAR(GETDATE())-1) " +
                    "AND OCCURRED_MONTH > 6 OR ISNULL(PROJECT_ID, '') = '' AND STATUS = 30 AND OCCURRED_YEAR = IIF(MONTH(GETDATE()) > 6, YEAR(GETDATE())+1, YEAR(GETDATE())) AND OCCURRED_MONTH < 6 GROUP BY PROJECT_ID)C " +
                    "ON p.PROJECT_ID = C.PROJECT_ID LEFT JOIN(SELECT PROJECT_ID, SUM(PAYMENT_TRANSFER) - SUM(RETENTION_PAYMENT) AS AP " +
                    "FROM PLAN_ESTIMATION_FORM WHERE ISNULL(INDIRECT_COST_TYPE, '') = '' GROUP BY PROJECT_ID)D ON p.PROJECT_ID = D.PROJECT_ID LEFT JOIN(SELECT PROJECT_ID, SUM(pop.AMOUNT) AS MinusItem FROM PLAN_ESTIMATION_FORM pef " +
                    "LEFT JOIN PLAN_OTHER_PAYMENT pop ON pef.EST_FORM_ID = pop.EST_FORM_ID WHERE pop.TYPE  in ('A', 'B', 'C', 'F') GROUP BY PROJECT_ID)E ON p.PROJECT_ID = E.PROJECT_ID " +
                    "LEFT JOIN (SELECT PROJECT_ID, COST AS MACost FROM PLAN_INDIRECT_COST WHERE FIELD_ID = '01_MACost')F on p.PROJECT_ID = F.PROJECT_ID " +
                    "LEFT JOIN (SELECT PROJECT_ID, COST AS SalesCost FROM PLAN_INDIRECT_COST WHERE FIELD_ID = '02_SalesCost')G on p.PROJECT_ID = G.PROJECT_ID " +
                    "LEFT JOIN (SELECT PROJECT_ID, SUM(AMOUNT) AS SiteCost FROM PLAN_SITE_BUDGET GROUP BY PROJECT_ID)H on p.PROJECT_ID = H.PROJECT_ID " +
                    "LEFT JOIN (SELECT PROJECT_ID, COST AS CompanyCost FROM PLAN_INDIRECT_COST WHERE FIELD_ID = '03_CompanyCost')I on p.PROJECT_ID = I.PROJECT_ID " +
                    "LEFT JOIN (SELECT PROJECT_ID, SUM(AMOUNT) AS SiteCostPaid FROM FIN_EXPENSE_FORM f LEFT JOIN FIN_EXPENSE_ITEM fi ON f.EXP_FORM_ID = fi.EXP_FORM_ID WHERE PROJECT_ID IS NOT NULL AND PROJECT_ID <> ''  AND STATUS = 30 GROUP BY PROJECT_ID)J on p.PROJECT_ID = J.PROJECT_ID " +
                    "LEFT JOIN (SELECT PROJECT_ID, ISNULL(SUM(PAYMENT_TRANSFER), 0) - ISNULL(SUM(RETENTION_PAYMENT), 0) AS MACostPaid " +
                    "FROM PLAN_ESTIMATION_FORM WHERE ISNULL(INDIRECT_COST_TYPE, '') = 'M' GROUP BY PROJECT_ID)K ON p.PROJECT_ID = K.PROJECT_ID " +
                    "LEFT JOIN (SELECT PROJECT_ID, ISNULL(SUM(PAYMENT_TRANSFER),0) - ISNULL(SUM(RETENTION_PAYMENT),0) AS OtherCostPaid " +
                    "FROM PLAN_ESTIMATION_FORM WHERE ISNULL(INDIRECT_COST_TYPE, '') = 'O' GROUP BY PROJECT_ID)M ON p.PROJECT_ID = M.PROJECT_ID " +
                    "GROUP BY p.PROJECT_ID, p.PROJECT_NAME, B.directCost, C.AR, D.AP, E.MinusItem, F.MACost, G.SalesCost, I.CompanyCost, H.SiteCost,J.SiteCostPaid,K.MACostPaid, M.OtherCostPaid)a)  AS futureCashFlow)cash ").FirstOrDefault();
            }
            return lstAmount;
        }
        public List<CreditNote> getCreditNoteById(string projectid, string formid)
        {
            logger.Debug("get credit note by form id, and form id =" + formid);
            List<CreditNote> lstItem = new List<CreditNote>();
            using (var context = new topmepEntities())
            {
                //條件篩選
                lstItem = context.Database.SqlQuery<CreditNote>("SELECT note.*, pi.TYPE AS INVOICE_TYPE, S.REGISTER_ADDRESS, S.COMPANY_ID FROM (SELECT I.*, P.OWNER_NAME  FROM PLAN_INVOICE I, TND_PROJECT P WHERE I.CONTRACT_ID = P.PROJECT_ID AND TYPE = '折讓單' " +
                    "AND EST_FORM_ID =@formid)note LEFT JOIN (SELECT * FROM PLAN_INVOICE WHERE CONTRACT_ID =@projectid AND TYPE <> '折讓單')pi ON note.INVOICE_NUMBER = pi.INVOICE_NUMBER LEFT JOIN TND_SUPPLIER S ON note.OWNER_NAME = S.COMPANY_NAME ", new SqlParameter("formid", formid), new SqlParameter("projectid", projectid)).ToList();
            }
            return lstItem;
        }

    }
}

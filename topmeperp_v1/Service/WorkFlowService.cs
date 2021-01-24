using log4net;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    //public enum ExecuteState
    //{
    //    Idle,  //停滯
    //    Running, //執行中
    //    Complete, //執行完成
    //    Fail, //錯誤
    //    Cancel, //取消
    //    Jump //跳過此Step
    //}
    //public interface IFlowContext
    //{
    //    System.Collections.Generic.Dictionary Parameters { get; }
    //    T GetParameter(stringkey);
    //    voidSetParameter(stringkey, T value);
    //}
    //public interface IFlowStep
    //{
    //    //當Step執行完畢後觸發
    //    event EventHandler ExecuteComplete;
    //    //當Step執行期間產生未補捉例外時,放置於此處
    //    Exception FailException { get; }
    //    //所有Flow Step共用的Context
    //    IFlowContext Context { get; set; }
    //    //Step Name, 必要欄位
    //    string StepName { get; }
    //    //Step執行狀態
    //    ExecuteState State { get; }
    //    //下一個要執行的Step Name,如未指定則依序執行
    //    string NextFlow { get; set; }
    //    //執行
    //    void Execute();
    //    //確認完成
    //    void Complete();
    //    //要求取消
    //    void Cancel();
    //    //跳過此Step
    //    void Jump();
    //    //Step執行失敗
    //    void Fail(Exception ex);
    //}
    public class WorkFlowService : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public WF_PROCESS process { get; set; }
        public List<WF_PROCESS_ACTIVITY> activitys { get; set; }

        public ExpenseTask task = null;
        protected string sql4Request = "SELECT * FROM WF_PROCESS_REQUEST WHERE DATA_KEY=@dataKey";
        protected string sql4Task = "SELECT * from WF_PORCESS_TASK WHERE RID=@rid ORDER BY SEQ_ID";
        protected string sqlDeptInfo = @"SELECT R.REQ_USER_ID REQ_USER_ID,U.USER_NAME REQ_USER_NAME,
                        U.DEP_CODE DEPT_CODE,D.DEPT_NAME DEPT_NAME,
                        D.MANAGER,(SELECT TOP 1 USER_NAME FROM SYS_USER  WHERE USER_ID=D.MANAGER) MANAGER_NAME
                         FROM WF_PROCESS_REQUEST r LEFT JOIN 
                        SYS_USER u  ON r.REQ_USER_ID=u.USER_ID LEFT OUTER JOIN
                        ENT_DEPARTMENT D ON u.DEP_CODE=D.DEPT_CODE
                        WHERE R.DATA_KEY=@datakey";
        protected SYS_USER user = null;
        public string statusChange = null;//More 、 Done、Fail
        public string Message = null;
        public enum TaskStatus
        {
            O, // Open
            R, //Running
            C, //Complete執行完成
            RJ, //Reject 拒絕 
            CAN, //Cancel 取消
            Jump //跳過此Step
        }

        public void getFlowAcivities(string flowkey)
        {
            using (var context = new topmepEntities())
            {
                string sql = "SELECT * FROM WF_PROCESS WHERE PROCESS_CODE=@processCode";
                logger.Debug("get process PROCESS_CODE=" + flowkey + ",sql=" + sql);
                process = context.WF_PROCESS.SqlQuery(sql, new SqlParameter("processCode", flowkey)).First();
                sql = "SELECT * FROM WF_PROCESS_ACTIVITY WHERE PID=@pid";
                logger.Debug("get activities pid=" + process.PID + ",sql=" + sql);
                activitys = context.WF_PROCESS_ACTIVITY.SqlQuery(sql, new SqlParameter("pid", process.PID)).ToList();
            }
        }
        //送審
        public void Send(SYS_USER u)
        {
            user = u;
            processTask();
        }
        //退件
        public void Reject(SYS_USER u, string reason)
        {
            user = u;
            RejectTask(reason);
        }
        //中止
        public void CancelRequest(SYS_USER u)
        {
            user = u;
            CancelRequest();
        }
        protected void processTask()
        {
            int idx = 0;
            for (idx = 0; idx < task.ProcessTask.Count; idx++)
            {
                switch (task.ProcessTask[idx].STATUS)
                {
                    case "O":
                        //change request status
                        if (idx + 1 < task.ProcessTask.Count)
                        {
                            //Has Next Step
                            UpdateTask(task.ProcessTask[idx], task.ProcessTask[idx + 1]);
                            if (statusChange == null)
                            {
                                statusChange = "M";//More
                            }
                            return;
                        }
                        else
                        {
                            //No More Step
                            UpdateTask(task.ProcessTask[idx], null);
                            if (statusChange == null)
                            {
                                statusChange = "D";//Done
                            }
                        }
                        break;
                    case "D":
                        //skip task
                        logger.Debug("task id=" + task.ProcessTask[idx].ID);
                        break;
                }
            }
        }
        /// <summary>
        /// 更新現有任務狀態
        /// </summary>
        /// <param name="curTask"></param>
        /// <param name="nextTask"></param>
        protected void UpdateTask(WF_PORCESS_TASK curTask, WF_PORCESS_TASK nextTask)
        {
            string sql4Task = "UPDATE WF_PORCESS_TASK SET STATUS=@status,REMARK=@remark,MODIFY_USER_ID=@modifyUser,MODIFY_DATE=@modifyDate WHERE ID=@id";
            string sql4Request = "UPDATE WF_PROCESS_REQUEST SET CURENT_STATE=@state,MODIFY_USER_ID=@modifyUser,MODIFY_DATE=@modifyDate WHERE RID=@RID";
            using (var context = new topmepEntities())
            {
                try
                {
                    //Update Task State=Done
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("status", "D"));
                    if (null != curTask.REMARK)
                    {
                        parameters.Add(new SqlParameter("remark", curTask.REMARK));
                    }
                    else
                    {
                        parameters.Add(new SqlParameter("remark", DBNull.Value));
                    }
                    parameters.Add(new SqlParameter("modifyUser", user.USER_ID));
                    parameters.Add(new SqlParameter("modifyDate", DateTime.Now));
                    parameters.Add(new SqlParameter("id", curTask.ID));
                    //Change Request State
                    int i = context.Database.ExecuteSqlCommand(sql4Task, parameters.ToArray());
                    logger.Debug("i=" + i + "sql" + sql4Task + ",Id" + curTask.ID);
                    parameters = new List<SqlParameter>();
                    if (null != nextTask)
                    {
                        parameters.Add(new SqlParameter("state", nextTask.SEQ_ID));
                    }
                    else
                    {
                        parameters.Add(new SqlParameter("state", -1));
                    }
                    parameters.Add(new SqlParameter("modifyUser", user.USER_ID));
                    parameters.Add(new SqlParameter("modifyDate", DateTime.Now));
                    parameters.Add(new SqlParameter("RID", curTask.RID));

                    i = context.Database.ExecuteSqlCommand(sql4Request, parameters.ToArray());
                    logger.Debug("i=" + i + "sql" + sql4Task + ",Id" + curTask.ID);
                    Message = "處理成功";
                }
                catch (Exception ex)
                {
                    statusChange = "F";
                    Message = "處理失敗(" + ex.Message + ")";
                    logger.Error(ex.Message + ":" + ex.StackTrace);
                }
            }
        }
        /// <summary>
        /// 退件相關狀態處理
        /// </summary>
        protected void RejectTask(string reason)
        {
            string sql4Task = "UPDATE WF_PORCESS_TASK SET STATUS=@status,REMARK=@remark,MODIFY_USER_ID=@modifyUser,MODIFY_DATE=@modifyDate WHERE RID=@RID";
            string sql4Request = "UPDATE WF_PROCESS_REQUEST SET CURENT_STATE=@state,MODIFY_USER_ID=@modifyUser,MODIFY_DATE=@modifyDate WHERE RID=@RID";
            using (var context = new topmepEntities())
            {
                try
                {
                    //Update Task State roll back to "O"
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("status", "O"));
                    parameters.Add(new SqlParameter("remark", reason));

                    parameters.Add(new SqlParameter("modifyUser", user.USER_ID));
                    parameters.Add(new SqlParameter("modifyDate", DateTime.Now));
                    parameters.Add(new SqlParameter("RID", task.task.RID));
                    int i = context.Database.ExecuteSqlCommand(sql4Task, parameters.ToArray());
                    logger.Debug("i=" + i + "sql" + sql4Task + ",Id" + task.task.RID);

                    //Change Request State
                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("state", 1));
                    parameters.Add(new SqlParameter("modifyUser", user.USER_ID));
                    parameters.Add(new SqlParameter("modifyDate", DateTime.Now));
                    parameters.Add(new SqlParameter("RID", task.task.RID));
                    i = context.Database.ExecuteSqlCommand(sql4Request, parameters.ToArray());
                    logger.Debug("i=" + i + "sql" + sql4Task + ",Id" + task.task.ID);
                    statusChange = "D";
                    Message = "退件作業完成";
                }
                catch (Exception ex)
                {
                    statusChange = "F";
                    Message = "退件作業處理失敗(" + ex.Message + ")";
                    logger.Error(ex.Message + ":" + ex.StackTrace);
                }
            }
        }
        /// <summary>
        /// 取消表單作業流程
        /// </summary>
        protected void CancelRequest()
        {
            string sql4Task = "DELETE WF_PORCESS_TASK WHERE RID=@RID";
            string sql4Request = "DELETE WF_PROCESS_REQUEST WHERE RID=@RID";
            using (var context = new topmepEntities())
            {
                try
                {
                    //Update Task State roll back to "O"
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("RID", task.task.RID));
                    int i = context.Database.ExecuteSqlCommand(sql4Task, parameters.ToArray());
                    logger.Debug("i=" + i + "sql" + sql4Task + ",Id" + task.task.RID);
                    i = context.Database.ExecuteSqlCommand(sql4Request, parameters.ToArray());
                    logger.Debug("i=" + i + "sql" + sql4Task + ",Id" + task.task.ID);
                    statusChange = "D";
                    Message = "表單已取消";
                }
                catch (Exception ex)
                {
                    statusChange = "F";
                    Message = "取消作業處理失敗(" + ex.Message + ")";
                    logger.Error(ex.Message + ":" + ex.StackTrace);
                }
            }
        }
        //取得表單與對應的流程資料
        public void getRequest(string datakey)
        {
            using (var context = new topmepEntities())
            {
                if (task == null)
                {
                    logger.Warn("task is null");
                    task = new ExpenseTask();
                }
                logger.Debug("get Request =" + datakey + ",sql=" + sql4Request);
                task.ProcessRequest = context.WF_PROCESS_REQUEST.SqlQuery(sql4Request, new SqlParameter("dataKey", datakey)).First();
                logger.Debug("get task rid=" + task.ProcessRequest.RID + ",sql=" + sql4Task);
                task.ProcessTask = context.WF_PORCESS_TASK.SqlQuery(sql4Task, new SqlParameter("rid", task.ProcessRequest.RID)).ToList();
            }
        }
        //建立對應的流程的相關任務
        protected void createFlow(SYS_USER u, string DataKey)
        {
            task = new ExpenseTask();
            //建立表單追蹤資料Index
            task.ProcessRequest = new WF_PROCESS_REQUEST();
            task.ProcessRequest.PID = process.PID;
            task.ProcessRequest.REQ_USER_ID = u.USER_ID;
            task.ProcessRequest.CREATE_DATE = DateTime.Now;
            task.ProcessRequest.DATA_KEY = DataKey;
            task.ProcessRequest.CURENT_STATE = 1;
            //建立簽核任務追蹤要項
            task.ProcessTask = new List<WF_PORCESS_TASK>();
            foreach (WF_PROCESS_ACTIVITY activity in activitys)
            {
                WF_PORCESS_TASK t = new WF_PORCESS_TASK();
                t.ACTIVITY_TYPE = activity.ACTIVITY_TYPE;
                t.CREATE_DATE = DateTime.Now;
                t.CREATE_USER_ID = u.USER_ID;
                t.NOTE = activity.ACTIVITY_NAME;
                t.STATUS = "O";//參考TaskStatus
                t.SEQ_ID = activity.SEQ_ID;
                t.DEP_CODE = activity.DEP_CODE;
                task.ProcessTask.Add(t);
            }
            using (var context = new topmepEntities())
            {
                context.WF_PROCESS_REQUEST.Add(task.ProcessRequest);
                int i = context.SaveChanges();
                foreach (WF_PORCESS_TASK t in task.ProcessTask)
                {
                    t.RID = task.ProcessRequest.RID;
                }
                context.WF_PORCESS_TASK.AddRange(task.ProcessTask);
                i = context.SaveChanges();
                logger.Debug("Create Task Records =" + i);
            }
        }
    }
    /// <summary>
    /// 公司費用申請控制流程
    /// </summary>
    public class Flow4CompanyExpense : WorkFlowService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public string FLOW_KEY = "EXP01";
        //處理SQL 預先填入專案代號,設定集合處理參數
        string sql = @"SELECT F.EXP_FORM_ID,F.PROJECT_ID,F.OCCURRED_YEAR,F.OCCURRED_MONTH,F.PAYEE,F.PAYMENT_DATE,F.REMARK REQ_DESC,F.REJECT_DESC REJECT_DESC,F.PAID_AMOUNT,
                        R.REQ_USER_ID,R.CURENT_STATE,R.PID,
						(SELECT TOP 1 MANAGER FROM ENT_DEPARTMENT WHERE DEPT_CODE=CT.DEP_CODE) MANAGER,
                        CT.* ,M.FORM_URL + METHOD_URL as FORM_URL
						FROM (SELECT fef.EXP_FORM_ID,fef.PROJECT_ID,fef.OCCURRED_YEAR,fef.OCCURRED_MONTH,PAYEE,fef.PAYMENT_DATE, 
                        fef.REMARK,fef.REJECT_DESC, fef.STATUS,SUM(fei.AMOUNT) AS PAID_AMOUNT 
                        FROM FIN_EXPENSE_FORM fef LEFT JOIN FIN_EXPENSE_ITEM fei ON fef.EXP_FORM_ID = fei.EXP_FORM_ID 
                        GROUP BY fef.EXP_FORM_ID,fef.PROJECT_ID,fef.OCCURRED_YEAR,fef.OCCURRED_MONTH,PAYEE,fef.STATUS,
                        fef.PAYMENT_DATE, fef.REMARK,fef.REJECT_DESC)F,WF_PROCESS_REQUEST R,
                        WF_PORCESS_TASK CT ,
		                (SELECT P.PID,A.SEQ_ID,FORM_URL,METHOD_URL  FROM WF_PROCESS P,WF_PROCESS_ACTIVITY A WHERE P.PID=A.PID ) M
                        WHERE F.EXP_FORM_ID= R.DATA_KEY AND R.RID=CT.RID AND R.CURENT_STATE=CT.SEQ_ID
						AND M.PID=R.PID AND M.SEQ_ID=R.CURENT_STATE";

        public Flow4CompanyExpense()
        {

        }
        /// <summary>
        /// 取得費用申請清單
        /// </summary>
        /// <param name="occurreddate"></param>
        /// <param name="subjectname"></param>
        /// <param name="expid"></param>
        /// <param name="projectid"></param>
        /// <returns></returns>
        public List<ExpenseFlowTask> getCompanyExpenseRequest(string occurreddate, string subjectname, string expid, string projectid, string status)
        {
            logger.Info("search expense form by " + occurreddate + ", 費用單編號 =" + expid + ", 項目名稱 =" + subjectname + ", 專案編號 =" + projectid);
            List<ExpenseFlowTask> lstForm = new List<ExpenseFlowTask>();
            using (var context = new topmepEntities())
            {
                var parameters = new List<SqlParameter>();
                //申請日期
                if (null != occurreddate && occurreddate != "")
                {
                    sql = sql + " AND F.PAYMENT_DATE=@occurreddate ";
                    parameters.Add(new SqlParameter("occurreddate", occurreddate));
                }
                //申請主旨內容
                if (null != subjectname && subjectname != "")
                {
                    sql = sql + " AND F.REMARK Like @subjectname ";
                    parameters.Add(new SqlParameter("subjectname", "'%" + subjectname + "%'"));
                }
                //申請單號
                if (null != expid && expid != "")
                {
                    sql = sql + " AND  F.EXP_FORM_ID Like @expid ";
                    parameters.Add(new SqlParameter("expid", "'%" + expid + "%'"));
                }
                //工地費用
                if (null != projectid && projectid != "")
                {
                    sql = sql + " AND  F.PROJECT_ID = @projectid ";
                    parameters.Add(new SqlParameter("projectid", projectid));
                }
                //表單狀態
                if (null != status && status != "")
                {
                    sql = sql + " AND  F.STATUS = @status ";
                    parameters.Add(new SqlParameter("status", status));
                }

                lstForm = context.Database.SqlQuery<ExpenseFlowTask>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get expense form count=" + lstForm.Count);

            return lstForm;
        }
        /// <summary>
        /// 取的申請程序表單與相關資料
        /// </summary>
        /// <param name="dataKey"></param>
        public void getTask(string dataKey)
        {
            task = new ExpenseTask();
            sql = sql + " AND R.DATA_KEY=@datakey";
            using (var context = new topmepEntities())
            {
                try
                {
                    //取得現有任務資訊
                    logger.Debug("sql=" + sql + ",Data Key=" + dataKey);
                    task.task = context.Database.SqlQuery<ExpenseFlowTask>(sql, new SqlParameter("datakey", dataKey)).First();
                    //取得申請者部門資料
                    logger.Debug("sqlDeptInfo=" + sqlDeptInfo + ",Data Key=" + dataKey);
                    task.DeptInfo = context.Database.SqlQuery<RequestUserDeptInfo>(sqlDeptInfo, new SqlParameter("datakey", dataKey)).First();

                }
                catch (Exception ex)
                {
                    logger.Warn("not task!! ex=" + ex.Message + "," + ex.StackTrace);
                }
            }
        }

        //表單送審後，啟動對應程序
        public void iniRequest(SYS_USER u, string DataKey)
        {
            getFlowAcivities(FLOW_KEY);
            createFlow(u, DataKey);
        }

        //送審
        public void Send(SYS_USER u, DateTime? paymentdate, string reason)
        {
            //STATUS  0   退件
            //STATUS  10  草稿
            //STATUS  20  審核中
            //STATUS  30  通過
            //STATUS  40  中止
            //
            logger.Debug("CompanyExpenseRequest Send" + task.task.ID);
            base.Send(u);
            if (statusChange != "F")
            {
                int staus = 10;
                if (statusChange == "M")
                {
                    staus = 20;
                }
                else if (statusChange == "D")
                {
                    staus = 30;
                }
                staus = updateForm(paymentdate, reason, staus);
            }
        }
        //更新資料庫資料
        protected int updateForm(DateTime? paymentdate, string reason, int staus)
        {
            string sql = "UPDATE FIN_EXPENSE_FORM SET STATUS=@status,PAYMENT_DATE=@paymentDate,REJECT_DESC=@rejectDesc WHERE EXP_FORM_ID=@formId";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("status", staus));
            if (null != paymentdate)
            {
                parameters.Add(new SqlParameter("paymentDate", paymentdate));
            }
            else
            {
                parameters.Add(new SqlParameter("paymentDate", DBNull.Value));
            }

            if (null != reason)
            {
                parameters.Add(new SqlParameter("rejectDesc", reason));
            }
            else
            {
                parameters.Add(new SqlParameter("rejectDesc", DBNull.Value));
            }
            parameters.Add(new SqlParameter("formId", task.task.EXP_FORM_ID));
            using (var context = new topmepEntities())
            {
                logger.Debug("Change CompanyExpenseRequest Status=" + task.task.EXP_FORM_ID + "," + staus);
                context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            }
            return staus;
        }

        //退件
        public void Reject(SYS_USER u, DateTime? paymentdate, string reason)
        {
            logger.Debug("CompanyExpenseRequest Reject:" + task.task.RID);
            base.Reject(u, reason);
            if (statusChange != "F")
            {
                updateForm(paymentdate, reason, 0);
            }
        }
        //中止
        public void Cancel(SYS_USER u)
        {
            user = u;
            logger.Info("USER :" + user.USER_ID + " Cancel :" + task.task.EXP_FORM_ID);
            base.CancelRequest(u);
            if (statusChange != "F")
            {
                string sql = "DELETE FIN_EXPENSE_ITEM WHERE EXP_FORM_ID=@formId;DELETE FIN_EXPENSE_FORM WHERE EXP_FORM_ID=@formId;";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("formId", task.task.EXP_FORM_ID));
                using (var context = new topmepEntities())
                {
                    logger.Debug("Cancel CompanyExpenseRequest Status=" + task.task.EXP_FORM_ID);
                    context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                }
            }
        }
    }
    /// <summary>
    /// 工地費用申請
    /// </summary>
    public class Flow4SiteExpense : Flow4CompanyExpense
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        //public new string FLOW_KEY = "EXP02";
        public Flow4SiteExpense()
        {
            base.FLOW_KEY = "EXP02";
        }
    }

    /// <summary>
    /// 廠商計價請款
    /// </summary>
    public class Flow4Estimation : WorkFlowService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public string FLOW_KEY = "EST01";
        //處理SQL 預先填入專案代號,設定集合處理參數
        string sql = @"
SELECT F.EST_FORM_ID,
(SELECT FORM_NAME  FROM PLAN_SUP_INQUIRY WHERE INQUIRY_FORM_ID=F.CONTRACT_ID) FORM_NAME,
(SELECT PROJECT_NAME FROM TND_PROJECT WHERE PROJECT_ID=F.PROJECT_ID) PROJECT_NAME,
F.PROJECT_ID,
F.PAYEE,
(SELECT COMPANY_NAME FROM TND_SUPPLIER WHERE SUPPLIER_ID=F.PAYEE) SUPPLIER_NAME ,
F.PAYMENT_DATE, 
F.PAID_AMOUNT,
F.REMARK REQ_DESC,
F.REJECT_DESC REJECT_DESC,
R.REQ_USER_ID,R.CURENT_STATE,R.PID,
   (SELECT TOP 1 MANAGER FROM ENT_DEPARTMENT WHERE DEPT_CODE=CT.DEP_CODE) MANAGER,
   CT.* ,M.FORM_URL + METHOD_URL as FORM_URL,
   P.PROCESS_CODE,
   P.PROCESS_NAME
FROM PLAN_ESTIMATION_FORM F,
WF_PROCESS_REQUEST R,
WF_PROCESS P,
 WF_PORCESS_TASK CT , 
   (SELECT P.PID,A.SEQ_ID,FORM_URL,METHOD_URL  FROM WF_PROCESS P,WF_PROCESS_ACTIVITY A WHERE P.PID=A.PID ) M
   WHERE F.EST_FORM_ID= R.DATA_KEY AND R.RID=CT.RID AND R.CURENT_STATE=CT.SEQ_ID
AND M.PID=R.PID AND M.SEQ_ID=R.CURENT_STATE  
AND P.PID=R.PID
";
        public Flow4Estimation()
        {
            //base.FLOW_KEY = "EST01";
        }
        /// <summary>
        /// 取得廠商計價請款清單
        /// </summary>
        /// <param name="contractid"></param>
        /// <param name="payee"></param>
        /// <param name="estid"></param>
        /// <param name="projectid"></param>
        /// <returns></returns>
        public List<ExpenseFlowTask> getEstimationFormRequest(string contractid, string payee, string estid, string projectid, string status)
        {
            logger.Info("search est form by 計價單編號 =" + estid + ", 受款人 =" + payee + ", 專案編號 =" + projectid);
            List<ExpenseFlowTask> lstForm = new List<ExpenseFlowTask>();
            using (var context = new topmepEntities())
            {
                var parameters = new List<SqlParameter>();
                //廠商名稱(受款人)
                if (null != payee && payee != "")
                {
                    sql = sql + " AND PAYEE Like @payee ";
                    parameters.Add(new SqlParameter("payee", "%" + payee + "%"));
                }
                //申請單號
                if (null != estid && estid != "")
                {
                    sql = sql + " AND  F.EST_FORM_ID Like @expid ";
                    parameters.Add(new SqlParameter("estid", "%" + estid + "%"));
                }
                //專案名稱
                if (null != projectid && projectid != "")
                {
                    sql = sql + " AND  F.PROJECT_ID = @projectid ";
                    parameters.Add(new SqlParameter("projectid", projectid));
                }
                //表單狀態
                if (null != status && status != "")
                {
                    sql = sql + " AND  F.STATUS = @status ";
                    parameters.Add(new SqlParameter("status", status));
                }
                logger.Debug("sql=" + sql);
                lstForm = context.Database.SqlQuery<ExpenseFlowTask>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get est form count=" + lstForm.Count);

            return lstForm;
        }
        /// <summary>
        /// 取的申請程序表單與相關資料
        /// </summary>
        /// <param name="dataKey"></param>
        public void getTask(string dataKey)
        {
            task = new ExpenseTask();
            sql = sql + " AND R.DATA_KEY=@datakey";
            using (var context = new topmepEntities())
            {
                try
                {
                    task.task = context.Database.SqlQuery<ExpenseFlowTask>(sql, new SqlParameter("datakey", dataKey)).First();
                    //取得申請者部門資料
                    logger.Debug("sqlDeptInfo=" + sqlDeptInfo + ",Data Key=" + dataKey);
                    task.DeptInfo = context.Database.SqlQuery<RequestUserDeptInfo>(sqlDeptInfo, new SqlParameter("datakey", dataKey)).First();
                }
                catch (Exception ex)
                {
                    logger.Warn("not task!! ex=" + ex.Message + "," + ex.StackTrace);
                }
            }
        }

        //表單送審後，啟動對應程序
        public void iniRequest(SYS_USER u, string DataKey)
        {
            getFlowAcivities(FLOW_KEY);
            task = new ExpenseTask();
            //建立表單追蹤資料Index
            task.ProcessRequest = new WF_PROCESS_REQUEST();
            task.ProcessRequest.PID = process.PID;
            task.ProcessRequest.REQ_USER_ID = u.USER_ID;
            task.ProcessRequest.CREATE_DATE = DateTime.Now;
            task.ProcessRequest.DATA_KEY = DataKey;
            task.ProcessRequest.CURENT_STATE = 1;
            //建立簽核任務追蹤要項
            task.ProcessTask = new List<WF_PORCESS_TASK>();
            foreach (WF_PROCESS_ACTIVITY activity in activitys)
            {
                WF_PORCESS_TASK t = new WF_PORCESS_TASK();
                t.ACTIVITY_TYPE = activity.ACTIVITY_TYPE;
                t.CREATE_DATE = DateTime.Now;
                t.CREATE_USER_ID = u.USER_ID;
                t.NOTE = activity.ACTIVITY_NAME;
                t.STATUS = "O";//參考TaskStatus
                t.SEQ_ID = activity.SEQ_ID;
                t.DEP_CODE = activity.DEP_CODE;
                task.ProcessTask.Add(t);
            }
            using (var context = new topmepEntities())
            {
                context.WF_PROCESS_REQUEST.Add(task.ProcessRequest);
                int i = context.SaveChanges();
                foreach (WF_PORCESS_TASK t in task.ProcessTask)
                {
                    //t.ID = DBNull.Value;
                    t.RID = task.ProcessRequest.RID;
                }
                context.WF_PORCESS_TASK.AddRange(task.ProcessTask);
                i = context.SaveChanges();
                logger.Debug("Create Task Records =" + i);
            }
        }
        //送審
        public void Send(SYS_USER u, DateTime? paymentdate, string reason, string payee, string remark)
        {
            //STATUS  0   退件
            //STATUS  10  草稿
            //STATUS  20  審核中
            //STATUS  30  通過
            //STATUS  40  中止
            //
            logger.Debug("EstimationFormRequest Send" + task.task.ID);
            base.Send(u);
            if (statusChange != "F")
            {
                int staus = 10;
                if (statusChange == "M")
                {
                    staus = 20;
                }
                else if (statusChange == "D")
                {
                    staus = 30;
                }
                staus = updateForm(paymentdate, reason, staus, payee, remark);
            }
        }
        //更新資料庫資料
        protected int updateForm(DateTime? paymentdate, string reason, int staus, string payee, string remark)
        {
            string sql = "UPDATE PLAN_ESTIMATION_FORM SET STATUS=@status, PAYMENT_DATE=@paymentDate, REJECT_DESC=@rejectDesc, PAYEE=@payee, REMARK=@remark WHERE EST_FORM_ID=@formId";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("status", staus));
            if (null != paymentdate)
            {
                parameters.Add(new SqlParameter("paymentDate", paymentdate));
            }
            else
            {
                parameters.Add(new SqlParameter("paymentDate", DBNull.Value));
            }
            if (null != reason)
            {
                parameters.Add(new SqlParameter("rejectDesc", reason));
            }
            else
            {
                parameters.Add(new SqlParameter("rejectDesc", DBNull.Value));
            }
            if (null != payee)
            {
                parameters.Add(new SqlParameter("payee", payee));
            }
            else
            {
                parameters.Add(new SqlParameter("payee", DBNull.Value));
            }
            if (null != remark)
            {
                parameters.Add(new SqlParameter("remark", remark));
            }
            else
            {
                parameters.Add(new SqlParameter("remark", DBNull.Value));
            }
            parameters.Add(new SqlParameter("formId", task.task.EST_FORM_ID));
            using (var context = new topmepEntities())
            {
                logger.Debug("Change EstimationFormRequest Status=" + task.task.EST_FORM_ID + "," + staus);
                context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            }
            if (staus == 30)
            {
                staus = addAccountFromEst(task.task.EST_FORM_ID, task.task.PAYEE, task.task.PAYMENT_DATE, task.task.PROJECT_ID, task.task.PAID_AMOUNT);
            }
            return staus;
        }
        //將廠商計價付款資料寫入帳款(暫時)
        public int addAccountFromEst(string formid, string payee, DateTime? paymentdate, string projectid, decimal Amt)
        {

            logger.Info("add plan account detail from est form,est form id =" + formid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                string sql = "INSERT INTO PLAN_ACCOUNT (PROJECT_ID, PAYEE, ACCOUNT_FORM_ID, AMOUNT_PAID, PAYMENT_DATE, ISDEBIT, ACCOUNT_TYPE) " +
                              "VALUES ('" + task.task.PROJECT_ID + "','" + task.task.PAYEE + "', '" + task.task.EST_FORM_ID + "'," + task.task.PAID_AMOUNT + ", CONVERT(datetime, REPLACE(REPLACE('" + task.task.PAYMENT_DATE + "', '上午', ''), '下午', '') +case when charindex('上午', '" + task.task.PAYMENT_DATE + "')> 0 then 'AM' when charindex('下午', '" + task.task.PAYMENT_DATE + "')> 0 then 'PM' end), 'N', 'P') ";
                logger.Info("sql =" + sql);
                i = context.Database.ExecuteSqlCommand(sql);
                return i;
            }
        }

        /*
        //將廠商計價付款資料寫入帳款
        public int addAccountFromEst(string formid, string contractid)
        {

            logger.Info("add plan account detail from est form,est form id =" + formid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                string sql = "INSERT INTO PLAN_ACCOUNT (PROJECT_ID, CONTRACT_ID, ACCOUNT_FORM_ID, AMOUNT_PAID, PAYMENT_DATE, ISDEBIT) " +
                                   "SELECT a.PROJECT_ID, a.CONTRACT_ID, a.EST_FORM_ID AS ACCOUNT_FORM_ID,ISNULL(advanceAmt2, ISNULL(retentionAmt2, estAmt2)) AS AMOUNT_PAID, paidDate2 AS PAYMENT_DATE, 'N' AS ISDEBIT FROM ( " +
                                   "SELECT ppt.*, ef.EST_FORM_ID, IIF(ef.advanceAmt < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_ADVANCE_CASH_RATIO / 100, ef.PAID_AMOUNT * USANCE_ADVANCE_CASH_RATIO / 100) AS advanceCash, " +
                                   "IIF(ef.advanceAmt < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_ADVANCE_1_RATIO / 100, ef.PAID_AMOUNT * USANCE_ADVANCE_1_RATIO / 100) AS advanceAmt1, " +
                                   "IIF(ef.advanceAmt < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_ADVANCE_2_RATIO / 100, ef.PAID_AMOUNT * USANCE_ADVANCE_2_RATIO / 100) AS advanceAmt2, " +
                                   "IIF(ef.RETENTION_PAYMENT < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_RETENTION_CASH_RATIO / 100, ef.PAID_AMOUNT * USANCE_RETENTION_CASH_RATIO / 100) AS retentionCash, " +
                                   "IIF(ef.RETENTION_PAYMENT < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_RETENTION_1_RATIO / 100, ef.PAID_AMOUNT * USANCE_RETENTION_1_RATIO / 100) AS retentionAmt1, " +
                                   "IIF(ef.RETENTION_PAYMENT < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_RETENTION_2_RATIO / 100, ef.PAID_AMOUNT * USANCE_RETENTION_2_RATIO / 100) AS retentionAmt2, " +
                                   "IIF(ef.PAYMENT_TRANSFER > 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * PAYMENT_ESTIMATED_CASH_RATIO / 100, ef.PAID_AMOUNT * USANCE_GOODS_CASH_RATIO / 100) AS estCash, " +
                                   "IIF(ef.PAYMENT_TRANSFER > 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * PAYMENT_ESTIMATED_1_RATIO / 100, ef.PAID_AMOUNT * USANCE_GOODS_1_RATIO / 100) AS estAmt1, " +
                                   "IIF(ef.PAYMENT_TRANSFER > 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * PAYMENT_ESTIMATED_2_RATIO / 100, ef.PAID_AMOUNT * USANCE_GOODS_2_RATIO / 100) AS estAmt2, " +
                                   "IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3) AS dateBase1, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null) AS dateBase2, " +
                                   "IIF(DAY(ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3), IIF(IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null) is null, IIF(MONTH(ef.CREATE_DATE) > 11, CONVERT(varchar, YEAR(ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), IIF(DAY(ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null), IIF(MONTH(ef.CREATE_DATE) > 11, CONVERT(varchar, YEAR(ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), CONVERT(varchar, YEAR(ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)))), " +
                                   "CONVERT(varchar, YEAR(ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, DATE_3))) AS paidDateCash, " +
                                   "IIF(DAY(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3), IIF(IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null) is null, IIF(MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) > 11, " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), " +
                                   "IIF(DAY(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null), IIF(MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) > 11, " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)))), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, DATE_3))) AS paidDate1, " +
                                   "IIF(DAY(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3), IIF(IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null) is null, IIF(MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) > 11, " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), " +
                                   "IIF(DAY(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null), IIF(MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) > 11," +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)))), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, DATE_3))) AS paidDate2 " +
                                   "FROM PLAN_PAYMENT_TERMS ppt LEFT JOIN(SELECT ef.CONTRACT_ID, ef.EST_FORM_ID, ef.CREATE_ID, pop.advanceAmt, ef.RETENTION_PAYMENT, ef.PAID_AMOUNT, ef.PAYMENT_TRANSFER, ef.CREATE_DATE FROM PLAN_ESTIMATION_FORM ef LEFT JOIN(SELECT EST_FORM_ID, SUM(AMOUNT) AS advanceAmt FROM PLAN_OTHER_PAYMENT WHERE TYPE IN('A', 'B', 'C') GROUP BY EST_FORM_ID HAVING EST_FORM_ID ='" + task.task.EST_FORM_ID + "')pop " +
                                   "ON ef.EST_FORM_ID = pop.EST_FORM_ID WHERE ef.EST_FORM_ID ='" + task.task.EST_FORM_ID + "')ef ON ppt.CONTRACT_ID = ef.CONTRACT_ID WHERE ppt.CONTRACT_ID ='" + task.task.CONTRACT_ID + "')a WHERE ISNULL(advanceAmt2, ISNULL(retentionAmt2, estAmt2)) IS NOT NULL UNION " +

                                   "SELECT a.PROJECT_ID, a.CONTRACT_ID, a.EST_FORM_ID,ISNULL(advanceCash, ISNULL(retentionCash, estCash))amtCash, paidDateCash, 'N' AS ISDEBIT FROM ( " +
                                   "SELECT ppt.*, ef.EST_FORM_ID, IIF(ef.advanceAmt < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_ADVANCE_CASH_RATIO / 100, ef.PAID_AMOUNT * USANCE_ADVANCE_CASH_RATIO / 100) AS advanceCash, " +
                                   "IIF(ef.advanceAmt < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_ADVANCE_1_RATIO / 100, ef.PAID_AMOUNT * USANCE_ADVANCE_1_RATIO / 100) AS advanceAmt1, " +
                                   "IIF(ef.advanceAmt < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_ADVANCE_2_RATIO / 100, ef.PAID_AMOUNT * USANCE_ADVANCE_2_RATIO / 100) AS advanceAmt2, " +
                                   "IIF(ef.RETENTION_PAYMENT < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_RETENTION_CASH_RATIO / 100, ef.PAID_AMOUNT * USANCE_RETENTION_CASH_RATIO / 100) AS retentionCash, " +
                                   "IIF(ef.RETENTION_PAYMENT < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_RETENTION_1_RATIO / 100, ef.PAID_AMOUNT * USANCE_RETENTION_1_RATIO / 100) AS retentionAmt1, " +
                                   "IIF(ef.RETENTION_PAYMENT < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_RETENTION_2_RATIO / 100, ef.PAID_AMOUNT * USANCE_RETENTION_2_RATIO / 100) AS retentionAmt2, " +
                                   "IIF(ef.PAYMENT_TRANSFER > 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * PAYMENT_ESTIMATED_CASH_RATIO / 100, ef.PAID_AMOUNT * USANCE_GOODS_CASH_RATIO / 100) AS estCash, " +
                                   "IIF(ef.PAYMENT_TRANSFER > 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * PAYMENT_ESTIMATED_1_RATIO / 100, ef.PAID_AMOUNT * USANCE_GOODS_1_RATIO / 100) AS estAmt1, " +
                                   "IIF(ef.PAYMENT_TRANSFER > 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * PAYMENT_ESTIMATED_2_RATIO / 100, ef.PAID_AMOUNT * USANCE_GOODS_2_RATIO / 100) AS estAmt2, " +
                                   "IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3) AS dateBase1, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null) AS dateBase2, " +
                                   "IIF(DAY(ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3), IIF(IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null) is null, IIF(MONTH(ef.CREATE_DATE) > 11, CONVERT(varchar, YEAR(ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), IIF(DAY(ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null), IIF(MONTH(ef.CREATE_DATE) > 11, CONVERT(varchar, YEAR(ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), CONVERT(varchar, YEAR(ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)))), " +
                                   "CONVERT(varchar, YEAR(ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, DATE_3))) AS paidDateCash, " +
                                   "IIF(DAY(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3), IIF(IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null) is null, IIF(MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) > 11, " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), " +
                                   "IIF(DAY(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null), IIF(MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) > 11, " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)))), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, DATE_3))) AS paidDate1, " +
                                   "IIF(DAY(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3), IIF(IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null) is null, IIF(MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) > 11, " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), " +
                                   "IIF(DAY(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null), IIF(MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) > 11," +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)))), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, DATE_3))) AS paidDate2 " +
                                   "FROM PLAN_PAYMENT_TERMS ppt LEFT JOIN(SELECT ef.CONTRACT_ID, ef.EST_FORM_ID,ef.CREATE_ID, pop.advanceAmt, ef.RETENTION_PAYMENT, ef.PAID_AMOUNT, ef.PAYMENT_TRANSFER, ef.CREATE_DATE FROM PLAN_ESTIMATION_FORM ef LEFT JOIN(SELECT EST_FORM_ID, SUM(AMOUNT) AS advanceAmt FROM PLAN_OTHER_PAYMENT WHERE TYPE IN('A', 'B', 'C') GROUP BY EST_FORM_ID HAVING EST_FORM_ID ='" + task.task.EST_FORM_ID + "')pop " +
                                   "ON ef.EST_FORM_ID = pop.EST_FORM_ID WHERE ef.EST_FORM_ID ='" + task.task.EST_FORM_ID + "')ef ON ppt.CONTRACT_ID = ef.CONTRACT_ID WHERE ppt.CONTRACT_ID ='" + task.task.CONTRACT_ID + "')a WHERE ISNULL(advanceCash, ISNULL(retentionCash, estCash)) IS NOT NULL UNION " +

                                   "SELECT a.PROJECT_ID, a.CONTRACT_ID, a.EST_FORM_ID,ISNULL(advanceAmt1, ISNULL(retentionAmt1, estAmt1))amt1, paidDate1, 'N' AS ISDEBIT FROM ( " +
                                   "SELECT ppt.*, ef.EST_FORM_ID, IIF(ef.advanceAmt < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_ADVANCE_CASH_RATIO / 100, ef.PAID_AMOUNT * USANCE_ADVANCE_CASH_RATIO / 100) AS advanceCash, " +
                                   "IIF(ef.advanceAmt < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_ADVANCE_1_RATIO / 100, ef.PAID_AMOUNT * USANCE_ADVANCE_1_RATIO / 100) AS advanceAmt1, " +
                                   "IIF(ef.advanceAmt < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_ADVANCE_2_RATIO / 100, ef.PAID_AMOUNT * USANCE_ADVANCE_2_RATIO / 100) AS advanceAmt2, " +
                                   "IIF(ef.RETENTION_PAYMENT < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_RETENTION_CASH_RATIO / 100, ef.PAID_AMOUNT * USANCE_RETENTION_CASH_RATIO / 100) AS retentionCash, " +
                                   "IIF(ef.RETENTION_PAYMENT < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_RETENTION_1_RATIO / 100, ef.PAID_AMOUNT * USANCE_RETENTION_1_RATIO / 100) AS retentionAmt1, " +
                                   "IIF(ef.RETENTION_PAYMENT < 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * ppt.PAYMENT_RETENTION_2_RATIO / 100, ef.PAID_AMOUNT * USANCE_RETENTION_2_RATIO / 100) AS retentionAmt2, " +
                                   "IIF(ef.PAYMENT_TRANSFER > 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * PAYMENT_ESTIMATED_CASH_RATIO / 100, ef.PAID_AMOUNT * USANCE_GOODS_CASH_RATIO / 100) AS estCash, " +
                                   "IIF(ef.PAYMENT_TRANSFER > 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * PAYMENT_ESTIMATED_1_RATIO / 100, ef.PAID_AMOUNT * USANCE_GOODS_1_RATIO / 100) AS estAmt1, " +
                                   "IIF(ef.PAYMENT_TRANSFER > 0 AND ppt.PAYMENT_TERMS = 'P', ef.PAID_AMOUNT * PAYMENT_ESTIMATED_2_RATIO / 100, ef.PAID_AMOUNT * USANCE_GOODS_2_RATIO / 100) AS estAmt2, " +
                                   "IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3) AS dateBase1, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null) AS dateBase2, " +
                                   "IIF(DAY(ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3), IIF(IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null) is null, IIF(MONTH(ef.CREATE_DATE) > 11, CONVERT(varchar, YEAR(ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), IIF(DAY(ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null), IIF(MONTH(ef.CREATE_DATE) > 11, CONVERT(varchar, YEAR(ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), CONVERT(varchar, YEAR(ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)))), " +
                                   "CONVERT(varchar, YEAR(ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, DATE_3))) AS paidDateCash, " +
                                   "IIF(DAY(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3), IIF(IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null) is null, IIF(MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) > 11, " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), " +
                                   "IIF(DAY(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null), IIF(MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) > 11, " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)))), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE1, USANCE_UP_TO_U_DATE1) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, DATE_3))) AS paidDate1, " +
                                   "IIF(DAY(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3), IIF(IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null) is null, IIF(MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) > 11, " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), " +
                                   "IIF(DAY(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) > IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, null), IIF(MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) > 11," +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) + 1) + '/' + '01' + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE) + 1) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3))), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_1, DATE_3)))), " +
                                   "CONVERT(varchar, YEAR(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, MONTH(IIF(ppt.PAYMENT_TERMS = 'P', PAYMENT_UP_TO_U_DATE2, USANCE_UP_TO_U_DATE2) + ef.CREATE_DATE)) + '/' + CONVERT(varchar, IIF(ppt.PAYMENT_FREQUENCY = 'O', DATE_2, DATE_3))) AS paidDate2 " +
                                   "FROM PLAN_PAYMENT_TERMS ppt LEFT JOIN(SELECT ef.CONTRACT_ID, ef.EST_FORM_ID,ef.CREATE_ID, pop.advanceAmt, ef.RETENTION_PAYMENT, ef.PAID_AMOUNT, ef.PAYMENT_TRANSFER, ef.CREATE_DATE FROM PLAN_ESTIMATION_FORM ef LEFT JOIN(SELECT EST_FORM_ID, SUM(AMOUNT) AS advanceAmt FROM PLAN_OTHER_PAYMENT WHERE TYPE IN('A', 'B', 'C') GROUP BY EST_FORM_ID HAVING EST_FORM_ID ='" + task.task.EST_FORM_ID + "')pop " +
                                   "ON ef.EST_FORM_ID = pop.EST_FORM_ID WHERE ef.EST_FORM_ID ='" + task.task.EST_FORM_ID + "')ef ON ppt.CONTRACT_ID = ef.CONTRACT_ID WHERE ppt.CONTRACT_ID ='" + task.task.CONTRACT_ID + "' )a WHERE ISNULL(advanceAmt1, ISNULL(retentionAmt1, estAmt1)) IS NOT NULL ";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("formid", formid));
                parameters.Add(new SqlParameter("contractid", contractid));
                i = context.Database.ExecuteSqlCommand(sql);
                return i;
            }
        }
        */
        //退件
        public void Reject(SYS_USER u, DateTime? paymentdate, string reason, string payee)
        {
            logger.Debug("EstimationFormRequest Reject:" + task.task.RID);
            base.Reject(u, reason);
            if (statusChange != "F")
            {
                updateForm(paymentdate, reason, 0, payee, null);
            }
        }
        //中止
        public void Cancel(SYS_USER u)
        {
            user = u;
            logger.Info("USER :" + user.USER_ID + " Cancel :" + task.task.EST_FORM_ID);
            base.CancelRequest(u);
            if (statusChange != "F")
            {
                string sql = "DELETE PLAN_ESTIMATION_FORM WHERE EST_FORM_ID=@formId;DELETE PLAN_ESTIMATION_FORM WHERE EST_FORM_ID=@formId;";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("formId", task.task.EST_FORM_ID));
                using (var context = new topmepEntities())
                {
                    logger.Debug("Cancel EstimationFormRequest Status=" + task.task.EST_FORM_ID);
                    context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                }
            }
        }
    }
    /// <summary>
    /// 估驗單審核作業
    /// </summary>
    public class Flow4EstimationForm : WorkFlowService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public string FLOW_KEY = "EST02";

        //處理SQL 預先填入專案代號,設定集合處理參數
        string sql = @"
SELECT F.EST_FORM_ID EXP_FORM_ID,F.EST_FORM_ID ,
(SELECT Top 1 FORM_NAME FROM PLAN_SUP_INQUIRY C WHERE C.INQUIRY_FORM_ID=F.CONTRACT_ID) FORM_NAME,
(SELECT Top 1 PROJECT_NAME FROM TND_PROJECT WHERE PROJECT_ID=F.PROJECT_ID) PROJECT_NAME ,
F.PROJECT_ID,
(SELECT TOP 1 COMPANY_NAME FROM TND_SUPPLIER WHERE SUPPLIER_ID=F.PAYEE) PAYEE,
F.PAYMENT_DATE,ISNULL(F.PAID_AMOUNT,0) PAID_AMOUNT,F.REMARK REQ_DESC,F.REJECT_DESC REJECT_DESC,
R.REQ_USER_ID,R.CURENT_STATE,R.PID,
(SELECT TOP 1 MANAGER FROM ENT_DEPARTMENT WHERE DEPT_CODE=CT.DEP_CODE) MANAGER,
CT.* ,M.FORM_URL + METHOD_URL as FORM_URL
FROM PLAN_ESTIMATION_FORM F,WF_PROCESS_REQUEST R,
WF_PORCESS_TASK CT , 
(SELECT P.PID,A.SEQ_ID,FORM_URL,METHOD_URL  FROM WF_PROCESS P,WF_PROCESS_ACTIVITY A WHERE P.PID=A.PID ) M
WHERE F.EST_FORM_ID= R.DATA_KEY AND R.RID=CT.RID AND R.CURENT_STATE=CT.SEQ_ID
AND M.PID=R.PID AND M.SEQ_ID=R.CURENT_STATE  
";

        public Flow4EstimationForm()
        {
        }
        /// <summary>
        /// 取得廠商計價請款清單
        /// </summary>
        public List<ExpenseFlowTask> getEstimationFormRequest(string contractid, string payee, string estid, string projectid, string status)
        {
            logger.Info("search est form by 計價單編號 =" + estid + ", 受款人 =" + payee + ", 專案編號 =" + projectid);
            List<ExpenseFlowTask> lstForm = new List<ExpenseFlowTask>();
            using (var context = new topmepEntities())
            {
                var parameters = new List<SqlParameter>();
                //廠商名稱(受款人)
                if (null != payee && payee != "")
                {
                    sql = sql + " AND PAYEE Like @payee ";
                    parameters.Add(new SqlParameter("payee", "%" + payee + "%"));
                }
                //申請單號
                if (null != estid && estid != "")
                {
                    sql = sql + " AND  F.EST_FORM_ID Like @expid ";
                    parameters.Add(new SqlParameter("estid", "%" + estid + "%"));
                }
                //專案名稱
                if (null != projectid && projectid != "")
                {
                    sql = sql + " AND  F.PROJECT_ID = @projectid ";
                    parameters.Add(new SqlParameter("projectid", projectid));
                }
                //表單狀態
                if (null != status && status != "")
                {
                    sql = sql + " AND  F.STATUS = @status ";
                    parameters.Add(new SqlParameter("status", status));
                }
                //排除已經進入財行(審核通過之估驗詹
                sql = sql + " AND  R.PID = 5 ";

                logger.Debug("sql=" + sql);
                lstForm = context.Database.SqlQuery<ExpenseFlowTask>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get est form count=" + lstForm.Count);

            return lstForm;
        }
        /// <summary>
        /// 取的申請程序表單與相關資料
        /// </summary>
        /// <param name="dataKey"></param>
        public void getTask(string dataKey)
        {
            task = new ExpenseTask();
            sql = sql + " AND R.DATA_KEY=@datakey";
            using (var context = new topmepEntities())
            {
                try
                {
                    task.task = context.Database.SqlQuery<ExpenseFlowTask>(sql, new SqlParameter("datakey", dataKey)).First();
                    //取得申請者部門資料
                    logger.Debug("sqlDeptInfo=" + sqlDeptInfo + ",Data Key=" + dataKey);
                    task.DeptInfo = context.Database.SqlQuery<RequestUserDeptInfo>(sqlDeptInfo, new SqlParameter("datakey", dataKey)).First();
                }
                catch (Exception ex)
                {
                    logger.Warn("not task!! ex=" + ex.Message + "," + ex.StackTrace);
                }
            }
        }

        //表單送審後，啟動對應程序
        public void iniRequest(SYS_USER u, string DataKey)
        {
            getFlowAcivities(FLOW_KEY);
            task = new ExpenseTask();
            //建立表單追蹤資料Index
            task.ProcessRequest = new WF_PROCESS_REQUEST();
            task.ProcessRequest.PID = process.PID;
            task.ProcessRequest.REQ_USER_ID = u.USER_ID;
            task.ProcessRequest.CREATE_DATE = DateTime.Now;
            task.ProcessRequest.DATA_KEY = DataKey;
            task.ProcessRequest.CURENT_STATE = 1;
            //建立簽核任務追蹤要項
            task.ProcessTask = new List<WF_PORCESS_TASK>();
            foreach (WF_PROCESS_ACTIVITY activity in activitys)
            {
                WF_PORCESS_TASK t = new WF_PORCESS_TASK();
                t.ACTIVITY_TYPE = activity.ACTIVITY_TYPE;
                t.CREATE_DATE = DateTime.Now;
                t.CREATE_USER_ID = u.USER_ID;
                t.NOTE = activity.ACTIVITY_NAME;
                t.STATUS = "O";//參考TaskStatus
                t.SEQ_ID = activity.SEQ_ID;
                t.DEP_CODE = activity.DEP_CODE;
                task.ProcessTask.Add(t);
            }
            using (var context = new topmepEntities())
            {
                context.WF_PROCESS_REQUEST.Add(task.ProcessRequest);
                int i = context.SaveChanges();
                foreach (WF_PORCESS_TASK t in task.ProcessTask)
                {
                    //t.ID = DBNull.Value;
                    t.RID = task.ProcessRequest.RID;
                }
                context.WF_PORCESS_TASK.AddRange(task.ProcessTask);
                i = context.SaveChanges();
                logger.Debug("Create Task Records =" + i);
            }
        }
        //送審
        public void Send(SYS_USER u, DateTime? paymentdate, string reason, string payee, string remark)
        {
            //STATUS  0   退件
            //STATUS  10  草稿
            //STATUS  20  審核中
            //STATUS  30  通過
            //STATUS  40  中止
            //
            logger.Debug("EstimationFormRequest Send" + task.task.ID);
            if (null == task.ProcessTask)
            {
                logger.Info("Not Have Task:" + task.task.EXP_FORM_ID);
                getRequest(task.task.EXP_FORM_ID);
            }
            base.Send(u);
            if (statusChange != "F")
            {
                int staus = 10;
                if (statusChange == "M")
                {
                    staus = 20;
                }
                else if (statusChange == "D")
                {
                    staus = 30;
                    //建立付款程序
                    logger.Info("create payment process:"+ task.task.EXP_FORM_ID);
                    Flow4Estimation paymentService = new Flow4Estimation();
                    paymentService.iniRequest(u, task.task.EXP_FORM_ID);
                }
                staus = updateForm(paymentdate, reason, staus, payee, remark);
            }
        }
        //更新資料庫資料
        protected int updateForm(DateTime? paymentdate, string reason, int staus, string payee, string remark)
        {
            string sql = "UPDATE PLAN_ESTIMATION_FORM SET STATUS=@status, PAYMENT_DATE=@paymentDate, REJECT_DESC=@rejectDesc, PAYEE=@payee, REMARK=@remark WHERE EST_FORM_ID=@formId";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("status", staus));
            if (null != paymentdate)
            {
                parameters.Add(new SqlParameter("paymentDate", paymentdate));
            }
            else
            {
                parameters.Add(new SqlParameter("paymentDate", DBNull.Value));
            }
            if (null != reason)
            {
                parameters.Add(new SqlParameter("rejectDesc", reason));
            }
            else
            {
                parameters.Add(new SqlParameter("rejectDesc", DBNull.Value));
            }
            if (null != payee)
            {
                parameters.Add(new SqlParameter("payee", payee));
            }
            else
            {
                parameters.Add(new SqlParameter("payee", DBNull.Value));
            }
            if (null != remark)
            {
                parameters.Add(new SqlParameter("remark", remark));
            }
            else
            {
                parameters.Add(new SqlParameter("remark", DBNull.Value));
            }
            parameters.Add(new SqlParameter("formId", task.task.EST_FORM_ID));
            using (var context = new topmepEntities())
            {
                logger.Debug("Change EstimationFormRequest Status=" + task.task.EST_FORM_ID + "," + staus);
                context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            }
            if (staus == 30)
            {
                // staus = addAccountFromEst(task.task.EST_FORM_ID, task.task.PAYEE, task.task.PAYMENT_DATE, task.task.PROJECT_ID, task.task.PAID_AMOUNT);
            }
            return staus;
        }

        //退件
        public void Reject(SYS_USER u, DateTime? paymentdate, string reason, string payee)
        {
            logger.Debug("EstimationFormRequest Reject:" + task.task.RID);
            base.Reject(u, reason);
            if (statusChange != "F")
            {
                updateForm(paymentdate, reason, 0, payee, null);
                Cancel(u); //註銷程序資料，送審後再建立
            }
        }
        //中止
        public void Cancel(SYS_USER u)
        {
            user = u;
            logger.Info("USER :" + user.USER_ID + " Cancel :" + task.task.EST_FORM_ID);
            base.CancelRequest(u);
            if (statusChange != "F")
            {
                //中止估驗單
                string sql = "UPDATE PLAN_ESTIMATION_FORM SET STATUS = '0' WHERE EST_FORM_ID=@formId; ";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("formId", task.task.EST_FORM_ID));
                using (var context = new topmepEntities())
                {
                    logger.Debug("Cancel EstimationFormRequest Status=" + task.task.EST_FORM_ID + ",User ID=" + u.USER_ID);
                    context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                }
            }
        }
    }
    /// <summary>
    /// 成本異動流程
    /// </summary>
    public class Flow4CostChange : WorkFlowService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public string FLOW_KEY = "CCH01";
        public new CostChangeFormTask task;//CostChangeFormTask
        //處理SQL 預先填入專案代號,設定集合處理參數
        string sql = @"SELECT F.FORM_ID,F.PROJECT_ID,F.REJECT_DESC REJECT_DESC,F.REMARK_ITEM, F.REMARK_QTY,F.REMARK_PRICE,F.REMARK_OTHER,F.STATUS,
                        R.REQ_USER_ID,R.CURENT_STATE,R.PID,
						(SELECT TOP 1 MANAGER FROM ENT_DEPARTMENT WHERE DEPT_CODE=CT.DEP_CODE) MANAGER,
                        CT.* ,M.FORM_URL + METHOD_URL as FORM_URL
						FROM PLAN_COSTCHANGE_FORM F,WF_PROCESS_REQUEST R,
                        WF_PORCESS_TASK CT ,
		                (SELECT P.PID,A.SEQ_ID,FORM_URL,METHOD_URL  FROM WF_PROCESS P,WF_PROCESS_ACTIVITY A WHERE P.PID=A.PID ) M
                        WHERE F.FORM_ID= R.DATA_KEY AND R.RID=CT.RID AND R.CURENT_STATE=CT.SEQ_ID
						AND M.PID=R.PID AND M.SEQ_ID=R.CURENT_STATE";
        /// <summary>
        /// 取得成本異動單據
        /// </summary>
        public List<CostChangeTask> getCostChangeRequest(string projectId, string remark, string status)
        {
            logger.Info("search costchagefor form by " + projectId);
            List<CostChangeTask> lstForm = new List<CostChangeTask>();
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectId));
            if (null != remark && remark != "")
            {
                sql = sql + " AND (F.REMARK_ITEM Like @remark OR F.REMARK_QTY Like @remark OR F.REMARK_PRICE Like @remark OR F.REMARK_OTHER Like @remark) ";
                parameters.Add(new SqlParameter("remark", "%" + remark + "%"));
            }
            if (null != status && status != "")
            {
                sql = sql + " AND F.STATUS =@status";
                parameters.Add(new SqlParameter("status", status));
            }

            using (var context = new topmepEntities())
            {
                lstForm = context.Database.SqlQuery<CostChangeTask>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get expense form count=" + lstForm.Count);

            return lstForm;
        }
        /// <summary>
        /// 取得成本異動申請程序表單與相關資料
        /// </summary>
        /// <param name="dataKey"></param>
        public void getTask(string dataKey)
        {
            task = new CostChangeFormTask();
            sql = sql + " AND R.DATA_KEY=@datakey";
            using (var context = new topmepEntities())
            {
                try
                {
                    //取得現有任務資訊
                    logger.Debug("sql=" + sql + ",Data Key=" + dataKey);
                    task.task = context.Database.SqlQuery<ExpenseFlowTask>(sql, new SqlParameter("datakey", dataKey)).First();
                    //取得申請者部門資料
                    logger.Debug("sqlDeptInfo=" + sqlDeptInfo + ",Data Key=" + dataKey);
                    task.DeptInfo = context.Database.SqlQuery<RequestUserDeptInfo>(sqlDeptInfo, new SqlParameter("datakey", dataKey)).First();

                    logger.Debug("get WF_REQUEST rid=" + dataKey + ",sql=" + sql4Task);
                    task.ProcessRequest = context.WF_PROCESS_REQUEST.SqlQuery(sql4Request, new SqlParameter("dataKey", dataKey)).First();
                    logger.Debug("get task rid=" + task.ProcessRequest.RID + ",sql=" + sql4Task);
                    task.ProcessTask = context.WF_PORCESS_TASK.SqlQuery(sql4Task, new SqlParameter("rid", task.ProcessRequest.RID)).ToList();

                }
                catch (Exception ex)
                {
                    logger.Warn("not task!! ex=" + ex.Message + "," + ex.StackTrace);
                }
            }
        }

        //表單送審後，啟動對應程序
        public void iniRequest(SYS_USER u, string DataKey)
        {
            getFlowAcivities(FLOW_KEY);
            createFlow(u, DataKey);
        }

        //送審
        public void Send(SYS_USER u, string desc, string reason, string methodCode, DateTime? settlementDate)
        {
            logger.Debug("CostChange Request Send" + task.task.ID);
            base.task = task;
            base.Send(u);
            if (statusChange != "F")
            {
                int staus = 10;
                if (statusChange == "M")
                {
                    staus = 20;
                }
                else if (statusChange == "D")
                {
                    staus = 30;
                }
                staus = updateForm(reason, staus, reason, methodCode, settlementDate);
            }
        }
        //更新資料庫資料
        protected int updateForm(string desc, int staus, string reason, string method, DateTime? settlementDate)
        {
            string sql = @"UPDATE PLAN_COSTCHANGE_FORM SET STATUS=@status,REJECT_DESC=@rejectDesc,
                           REASON_CODE=@reason,METHOD_CODE=@Methodcode,SETTLEMENT_DATE=@settlementDate 
                           WHERE FORM_ID=@formId";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("formId", task.FormData.FORM_ID));
            parameters.Add(new SqlParameter("status", staus));

            if (null == desc)
            {
                sql = sql.Replace(",REJECT_DESC=@rejectDesc", "");
            }
            else
            {
                parameters.Add(new SqlParameter("rejectDesc", reason));
            }

            if (null == settlementDate)
            {
                sql = sql.Replace(",SETTLEMENT_DATE=@settlementDate", "");
            }
            else
            {
                parameters.Add(new SqlParameter("settlementDate", settlementDate));
            }
            if (null == reason)
            {
                sql = sql.Replace(",REASON_CODE=@reason", "");
            }
            else
            {
                parameters.Add(new SqlParameter("reason", reason));
            }

            if (null == method)
            {
                sql = sql.Replace(",METHOD_CODE=@Methodcode", "");
            }
            else
            {
                parameters.Add(new SqlParameter("Methodcode", method));
            }


            using (var context = new topmepEntities())
            {
                logger.Debug("Change CostChange Status=" + task.FormData.FORM_ID + "," + staus);
                context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            }
            return staus;
        }

        //退件
        public new void Reject(SYS_USER u, string reason)
        {
            logger.Debug("CostChangeRequest Reject:" + task.task.RID);
            base.task = task;
            base.Reject(u, reason);
            if (statusChange != "F")
            {
                updateForm(reason, 0, null, null, null);
            }
        }
        //中止
        public void Cancel(SYS_USER u)
        {
            user = u;
            logger.Info("USER :" + user.USER_ID + " Cancel :" + task.task.EXP_FORM_ID);
            base.task = task;
            base.CancelRequest(u);
            if (statusChange != "F")
            {
                string sql = "DELETE PLAN_COSTCHANGE_ITEM WHERE FORM_ID=@formId;DELETE PLAN_COSTCHANGE_FORM WHERE FORM_ID=@formId;";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("formId", task.FormData.FORM_ID));
                using (var context = new topmepEntities())
                {
                    logger.Debug("Cancel CostChangeRequest Status=" + task.FormData.FORM_ID);
                    context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                }
            }
        }
    }
}
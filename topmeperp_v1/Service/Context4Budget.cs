using log4net;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    public class Service4Budget : PurchaseFormService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        #region 工地費用
        /// <summary>
        /// 取得工地預算，依據專案與年度條件
        /// </summary>
        string sql4Buget = @"SELECT ROW_NUMBER() OVER(ORDER BY A.SUBJECT_ID) AS SUB_NO, A.*, G.HTOTAL FROM 
                (SELECT SUBJECT_NAME, FIN_SUBJECT_ID AS SUBJECT_ID, [01] As 'JAN', [02] As 'FEB', [03] As 'MAR', [04] As 'APR', [05] As 'MAY', [06] As 'JUN', 
                [07] As 'JUL', [08] As 'AUG', [09] As 'SEP', [10] As 'OCT', [11] As 'NOV', [12] As 'DEC' FROM
                (SELECT main.*, sub.AMOUNT, sub.BUDGET_YEAR, sub.BUDGET_MONTH FROM 
                (SELECT fs.FIN_SUBJECT_ID, fs.SUBJECT_NAME FROM FIN_SUBJECT fs WHERE fs.CATEGORY = '工地費用') main LEFT JOIN 
                (SELECT psb.BUDGET_MONTH, psb.AMOUNT, psb.BUDGET_YEAR, psb.SUBJECT_ID FROM PLAN_SITE_BUDGET psb 
                WHERE psb.PROJECT_ID = @projectid 
                AND (@targetYear is null or psb.BUDGET_YEAR = @targetYear) 
                AND (@yearSeq is null or psb.YEAR_SEQUENCE = @yearSeq))sub ON sub.SUBJECT_ID = main.FIN_SUBJECT_ID) As STable 
                PIVOT(SUM(AMOUNT) FOR BUDGET_MONTH IN([01], [02], [03], [04], [05], [06], [07], [08], [09], [10], [11], [12])) As PTable)A 
                LEFT JOIN 
                (SELECT SUBJECT_ID, ISNULL(SUM(psb.AMOUNT), 0) AS HTOTAL FROM PLAN_SITE_BUDGET psb 
                WHERE psb.PROJECT_ID = @projectid AND psb.BUDGET_YEAR < @targetYear OR psb.PROJECT_ID = @projectid 
                AND psb.BUDGET_YEAR = @targetYear GROUP BY SUBJECT_ID)G ON A.SUBJECT_ID = G.SUBJECT_ID  ";
        /// <summary>
        /// 取得工地費用彙整，與預算比較需要僅能透過年度件
        /// </summary>
        string sql4Expense = @"SELECT ROW_NUMBER() OVER(ORDER BY C.SUBJECT_ID) + 1 AS SUB_NO, C.* FROM 
                (SELECT OCCURRED_YEAR,SUBJECT_NAME, SUBJECT_ID, [01] As 'JAN', [02] As 'FEB', [03] As 'MAR', [04] As 'APR', [05] As 'MAY'
                , [06] As 'JUN', [07] As 'JUL', [08] As 'AUG', [09] As 'SEP', [10] As 'OCT', [11] As 'NOV', [12] As 'DEC' FROM
                (
                SELECT * FROM 
                (SELECT fs.FIN_SUBJECT_ID SUBJECT_ID, fs.SUBJECT_NAME FROM FIN_SUBJECT fs WHERE fs.CATEGORY = '工地費用') main
                LEFT OUTER JOIN 
                (SELECT ef.OCCURRED_YEAR,ef.OCCURRED_MONTH, ei.FIN_SUBJECT_ID, ei.AMOUNT FROM FIN_EXPENSE_ITEM ei  
                LEFT JOIN FIN_EXPENSE_FORM ef ON ei.EXP_FORM_ID = ef.EXP_FORM_ID 
                WHERE ef.PROJECT_ID = @projectid AND ef.OCCURRED_YEAR = @targetYear) Expen
                ON main.SUBJECT_ID = Expen.FIN_SUBJECT_ID 
                 ) As STable 
                PIVOT(SUM(AMOUNT) FOR OCCURRED_MONTH IN([01], [02], [03], [04], [05], [06], [07], [08], [09], [10], [11], [12])) As PTable) C ";
        //取得工地費用項目
        public List<FIN_SUBJECT> getSubjectOfExpense4Site()
        {
            List<FIN_SUBJECT> lstSubject = new List<FIN_SUBJECT>();
            using (var context = new topmepEntities())
            {
                lstSubject = context.Database.SqlQuery<FIN_SUBJECT>("SELECT * FROM FIN_SUBJECT WHERE CATEGORY = '工地費用' ORDER BY FIN_SUBJECT_ID; ").ToList();
                logger.Info("Get Subject of Operating Expense Count=" + lstSubject.Count);
            }
            return lstSubject;
        }
        //取得專案編號?
        public string getSiteBudgetById(string prjid)
        {
            string projectid = null;
            using (var context = new topmepEntities())
            {
                projectid = context.Database.SqlQuery<string>("SELECT DISTINCT PROJECT_ID FROM PLAN_SITE_BUDGET WHERE PROJECT_ID = @pid "
               , new SqlParameter("pid", prjid)).FirstOrDefault();
            }
            return projectid;
        }
        //刪除專案工地預算
        public int delSiteBudgetByProject(string projectid, string year)
        {
            logger.Info("remove all site budget by projectid =" + projectid + "and by year sequence =" + year);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete all PLAN_SITE_BUDGET by projectid =" + projectid + "and by year sequence =" + year);
                i = context.Database.ExecuteSqlCommand("DELETE FROM PLAN_SITE_BUDGET WHERE PROJECT_ID =@projectid AND YEAR_SEQUENCE =@year", new SqlParameter("@projectid", projectid), new SqlParameter("@year", year));
            }
            logger.Debug("delete PALN_SITE_BUDGET count=" + i);
            return i;
        }
        //匯入工地預算
        public int refreshSiteBudget(List<PLAN_SITE_BUDGET> items)
        {
            int i = 0;
            logger.Info("refreshSiteBudgetItem = " + items.Count);
            //2.將Excel 資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (PLAN_SITE_BUDGET item in items)
                {
                    context.PLAN_SITE_BUDGET.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add PLAN_SITE_BUDGET count =" + i);
            return i;
        }

        #endregion
        //取得專案工地費用預算之西元年
        #region 年度
        public int getSiteBudgetByYearSeq(string prjid, string yearSeq)
        {
            int year = 0;
            using (var context = new topmepEntities())
            {
                string sql = @"SELECT DISTINCT BUDGET_YEAR FROM PLAN_SITE_BUDGET WHERE PROJECT_ID = @pid AND YEAR_SEQUENCE = @yearSeq";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("pid", prjid));
                parameters.Add(new SqlParameter("yearSeq", yearSeq));
                year = context.Database.SqlQuery<int>(sql, parameters.ToArray()).FirstOrDefault();
            }
            return year;
        }
        #endregion
        //取得特定專案之工地費用預算年度
        public int getYearOfSiteExpenseById(string projectid, int seqYear)
        {
            int lstYear = 0;
            using (var context = new topmepEntities())
            {
                lstYear = context.Database.SqlQuery<int>("SELECT BUDGET_YEAR FROM PLAN_SITE_BUDGET WHERE PROJECT_ID = @projectid AND YEAR_SEQUENCE = @seqYear GROUP BY BUDGET_YEAR "
               , new SqlParameter("projectid", projectid), new SqlParameter("seqYear", seqYear)).FirstOrDefault();
            }
            return lstYear;
        }
        //取得特定專案之起始年度
        public int getFirstYearOfPlanById(string projectid)
        {
            int lstYear = 0;
            using (var context = new topmepEntities())
            {
                lstYear = context.Database.SqlQuery<int>("SELECT YEAR(CREATE_DATE) FROM PLAN_ITEM WHERE PROJECT_ID = @projectid GROUP BY YEAR(CREATE_DATE) "
               , new SqlParameter("projectid", projectid)).FirstOrDefault();
            }
            return lstYear;
        }

        //取得專案工地費用預算
        public List<ExpenseBudgetSummary> getBudget4ProjectBySeq(string projectid, string targetYear, string yearSeq)
        {
            List<ExpenseBudgetSummary> lstSiteBudget = new List<ExpenseBudgetSummary>();
            using (var context = new topmepEntities())
            {
                logger.Debug("sql=" + sql4Buget);
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", projectid));
                parameters.Add(new SqlParameter("targetYear", (object)targetYear ?? DBNull.Value));
                parameters.Add(new SqlParameter("yearSeq", (object)yearSeq ?? DBNull.Value));
                lstSiteBudget = context.Database.SqlQuery<ExpenseBudgetSummary>(sql4Buget, parameters.ToArray()).ToList();
            }
            return lstSiteBudget;
        }
        //取得特定期間工地費用預算與費用彙整
        public List<ExpenseBudgetSummary> getSiteExpenseSummaryByYear(string projectid, string targetYear)
        {
            List<ExpenseBudgetSummary> lstExpBudget = new List<ExpenseBudgetSummary>();
            using (var context = new topmepEntities())
            {
                logger.Info("sql = " + sql4Expense);
                lstExpBudget = context.Database.SqlQuery<ExpenseBudgetSummary>(sql4Expense, new SqlParameter("projectid", projectid), new SqlParameter("targetYear", targetYear)).ToList();
            }
            return lstExpBudget;
        }

        //取得特定工地每年預算
        public List<ExpenseBudgetSummary> getSiteBudgetPerYear(string projectid)
        {
            List<ExpenseBudgetSummary> lstSiteBudget = new List<ExpenseBudgetSummary>();
            using (var context = new topmepEntities())
            {
                string sql = @"SELECT CAST(YEAR_SEQUENCE AS VARCHAR) YEAR_SEQUENCE,CAST(BUDGET_YEAR as VARCHAR) BUDGET_YEAR ,SUM(AMOUNT) as TOTAL_BUDGET FROM PLAN_SITE_BUDGET 
                            WHERE PROJECT_ID=@projectid
                            GROUP BY BUDGET_YEAR,YEAR_SEQUENCE ORDER BY YEAR_SEQUENCE";
                logger.Debug("sql=" + sql + " ,projectId=" + projectid);
                lstSiteBudget = context.Database.SqlQuery<ExpenseBudgetSummary>(sql, new SqlParameter("projectid", projectid)).ToList();
            }
            return lstSiteBudget;
        }
        //取得特定專案工地費用每月執行總和
        public List<ExpensetFromOPByMonth> getSiteExpensetOfMonth(string projectid, int targetYear, int targetMonth, bool isCum)
        {
            List<ExpensetFromOPByMonth> lstExpense = new List<ExpensetFromOPByMonth>();
            using (var context = new topmepEntities())
            {
                if (isCum == true)
                {
                    lstExpense = context.Database.SqlQuery<ExpensetFromOPByMonth>("SELECT SUM(F.JAN) AS JAN, SUM(F.FEB) AS FEB, SUM(F.MAR) AS MAR, SUM(F.APR) AS APR, SUM(F.MAY) AS MAY, SUM(F.JUN) AS JUN, " +
                        "SUM(F.JUL) AS JUL, SUM(F.AUG) AS AUG, SUM(F.SEP) AS SEP, SUM(F.OCT) AS OCT, SUM(F.NOV) AS NOV, SUM(F.DEC) AS DEC, SUM(F.HTOTAL) AS HTOTAL FROM(SELECT C.*, E.HTOTAL " +
                        "FROM(SELECT SUBJECT_NAME, FIN_SUBJECT_ID, [01] As 'JAN', [02] As 'FEB', [03] As 'MAR', [04] As 'APR', [05] As 'MAY', [06] As 'JUN', [07] As 'JUL', [08] As 'AUG', [09] As 'SEP', [10] As 'OCT', [11] As 'NOV', [12] As 'DEC' " +
                        "FROM(SELECT B.OCCURRED_MONTH, fs.FIN_SUBJECT_ID, fs.SUBJECT_NAME, B.AMOUNT FROM FIN_SUBJECT fs LEFT JOIN(SELECT ef.OCCURRED_MONTH, ei.FIN_SUBJECT_ID, ei.AMOUNT FROM FIN_EXPENSE_ITEM ei " +
                        "LEFT JOIN FIN_EXPENSE_FORM ef ON ei.EXP_FORM_ID = ef.EXP_FORM_ID WHERE ef.PROJECT_ID = @projectid AND ef.OCCURRED_YEAR = @targetYear AND ef.OCCURRED_MONTH <= @targetMonth)B " +
                        "ON fs.FIN_SUBJECT_ID = B.FIN_SUBJECT_ID WHERE fs.CATEGORY = '工地費用') As STable " +
                        "PIVOT(SUM(AMOUNT) FOR OCCURRED_MONTH IN([01], [02], [03], [04], [05], [06], [07], [08], [09], [10], [11], [12])) As PTable)C LEFT JOIN(SELECT FIN_SUBJECT_ID, ISNULL(SUM(fei.AMOUNT), 0) AS HTOTAL " +
                        "FROM FIN_EXPENSE_ITEM fei LEFT JOIN FIN_EXPENSE_FORM fef ON fei.EXP_FORM_ID = fef.EXP_FORM_ID WHERE fef.PROJECT_ID = @projectid AND fef.OCCURRED_YEAR < @targetYear OR " +
                        "fef.PROJECT_ID = @projectid AND fef.OCCURRED_YEAR = @targetYear AND fef.OCCURRED_MONTH <= @targetMonth GROUP BY FIN_SUBJECT_ID)E ON C.FIN_SUBJECT_ID = E.FIN_SUBJECT_ID)F; "
                   , new SqlParameter("projectid", projectid), new SqlParameter("targetYear", targetYear), new SqlParameter("targetMonth", targetMonth)).ToList();
                }
                else
                {
                    lstExpense = context.Database.SqlQuery<ExpensetFromOPByMonth>("SELECT SUM(F.JAN) AS JAN, SUM(F.FEB) AS FEB, SUM(F.MAR) AS MAR, SUM(F.APR) AS APR, SUM(F.MAY) AS MAY, SUM(F.JUN) AS JUN, " +
                        "SUM(F.JUL) AS JUL, SUM(F.AUG) AS AUG, SUM(F.SEP) AS SEP, SUM(F.OCT) AS OCT, SUM(F.NOV) AS NOV, SUM(F.DEC) AS DEC, SUM(F.HTOTAL) AS HTOTAL FROM(SELECT  C.*, " +
                        "SUM(ISNULL(C.JAN, 0)) + SUM(ISNULL(C.FEB, 0)) + SUM(ISNULL(C.MAR, 0)) + SUM(ISNULL(C.APR, 0)) + SUM(ISNULL(C.MAY, 0)) + SUM(ISNULL(C.JUN, 0)) " +
                        "+ SUM(ISNULL(C.JUL, 0)) + SUM(ISNULL(C.AUG, 0)) + SUM(ISNULL(C.SEP, 0)) + SUM(ISNULL(C.OCT, 0)) + SUM(ISNULL(C.NOV, 0)) + SUM(ISNULL(C.DEC, 0)) AS HTOTAL " +
                        "FROM(SELECT SUBJECT_NAME, FIN_SUBJECT_ID, [01] As 'JAN', [02] As 'FEB', [03] As 'MAR', [04] As 'APR', [05] As 'MAY', [06] As 'JUN', [07] As 'JUL', [08] As 'AUG', [09] As 'SEP', [10] As 'OCT', [11] As 'NOV', [12] As 'DEC' " +
                        "FROM(SELECT B.OCCURRED_MONTH, fs.FIN_SUBJECT_ID, fs.SUBJECT_NAME, B.AMOUNT FROM FIN_SUBJECT fs LEFT JOIN(SELECT ef.OCCURRED_MONTH, ei.FIN_SUBJECT_ID, ei.AMOUNT FROM FIN_EXPENSE_ITEM ei " +
                        "LEFT JOIN FIN_EXPENSE_FORM ef ON ei.EXP_FORM_ID = ef.EXP_FORM_ID WHERE ef.PROJECT_ID = @projectid AND ef.OCCURRED_YEAR = @targetYear)B " +
                        "ON fs.FIN_SUBJECT_ID = B.FIN_SUBJECT_ID WHERE fs.CATEGORY = '工地費用') As STable " +
                        "PIVOT(SUM(AMOUNT) FOR OCCURRED_MONTH IN([01], [02], [03], [04], [05], [06], [07], [08], [09], [10], [11], [12])) As PTable)C " +
                        "GROUP BY C.SUBJECT_NAME, C.FIN_SUBJECT_ID, C.JAN, C.FEB, C.MAR, C.APR, C.MAY, C.JUN, C.JUL, C.AUG, C.SEP, C.OCT, C.NOV, C.DEC)F ; "
                        , new SqlParameter("projectid", projectid), new SqlParameter("targetYear", targetYear)).ToList();
                }
            }
            return lstExpense;
        }
        //取得特定專案之工地費用總預算金額
        public ExpenseBudgetSummary getSiteBudgetAmountById(string projectid, string budgetYear)
        {
            ExpenseBudgetSummary lstAmount = null;
            string sql = @"SELECT SUM(AMOUNT) AS TOTAL_BUDGET FROM PLAN_SITE_BUDGET 
                WHERE PROJECT_ID = @projectid  
                AND (@budgetYear IS NULL OR BUDGET_YEAR=@budgetYear)";
            using (var context = new topmepEntities())
            {
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", projectid));
                parameters.Add(new SqlParameter("budgetYear", (object)budgetYear ?? DBNull.Value));
                lstAmount = context.Database.SqlQuery<ExpenseBudgetSummary>(sql, parameters.ToArray()).FirstOrDefault();
            }
            return lstAmount;
        }

        //取得特定專案之工地費用總執行金額
        public ExpenseBudgetSummary getTotalSiteExpAmountById(string projectid, string targetYear)
        {
            string sql = @"SELECT SUM(ei.AMOUNT) AS CUM_YEAR_AMOUNT FROM FIN_EXPENSE_ITEM ei 
                    LEFT JOIN FIN_EXPENSE_FORM ef 
                    ON ei.EXP_FORM_ID = ef.EXP_FORM_ID WHERE ef.PROJECT_ID = @projectid 
                    AND ef.STATUS=@status
                    AND (@targetYear is Null OR ef.OCCURRED_YEAR = @targetYear)";
            ExpenseBudgetSummary lstExpAmount = null;
            using (var context = new topmepEntities())
            {
                logger.Debug("sql=" + sql + ",project_id=" + projectid + ",year=" + targetYear);
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("projectid", projectid));
                parameters.Add(new SqlParameter("status", "30"));
                parameters.Add(new SqlParameter("targetYear", (object)targetYear ?? DBNull.Value));
                lstExpAmount = context.Database.SqlQuery<ExpenseBudgetSummary>(sql, parameters.ToArray()).FirstOrDefault();
            }
            return lstExpAmount;
        }

        //取得特定年度之公司費用總預算金額
        public ExpenseBudgetSummary getTotalExpBudgetAmount(int year)
        {
            ExpenseBudgetSummary lstAmount = null;
            using (var context = new topmepEntities())
            {
                lstAmount = context.Database.SqlQuery<ExpenseBudgetSummary>("SELECT SUM(AMOUNT) AS TOTAL_BUDGET FROM FIN_EXPENSE_BUDGET WHERE BUDGET_YEAR = @year  "
               , new SqlParameter("year", year)).FirstOrDefault();
            }
            return lstAmount;
        }
    }
    /// <summary>
    /// 現金流量報表
    /// </summary>
    class Service4CashFlow : Service4Budget
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        //取得特定日期借款與還款
        public new List<LoanTranactionFunction> getLoanTranaction(string type, string paymentdate, string duringStart, string duringEnd)
        {
            logger.Debug("get loan tranaction by payment date=" + paymentdate + ",and type = " + type);
            List<LoanTranactionFunction> lstItem = null;
            string sql = @"SELECT t.*, l.IS_SUPPLIER 
                    FROM FIN_LOAN_TRANACTION t 
                    LEFT JOIN FIN_BANK_LOAN l ON t.BL_ID = l.BL_ID  
                    WHERE  t.REMARK NOT LIKE '%備償%' 
                    AND t.TRANSACTION_TYPE = @TransType $WhereCond
                    ";
            string strWhere = "";
            int transType = 1;
            //ISSUE : TYPE 何時為-1 對銀行借款，-1 錶現金流入
            if (type == "I")// type值為'I'表示有現金流入
            {
                transType = -1;
            }
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("TransType", transType));

            using (var context = new topmepEntities())
            {
                //條件篩選
                if (null != paymentdate && paymentdate != "")
                {
                    //just one day
                    strWhere = "AND  IIF(TRANSACTION_TYPE = 1,PAYBACK_DATE,EVENT_DATE) =@paymentdate ";
                    parameters.Add(new SqlParameter("paymentdate", paymentdate));
                    logger.Debug("paymentdate=" + paymentdate);
                }
                else if (null != duringStart && duringStart != "" && null != duringEnd && duringEnd != "")
                {
                    //for durateion
                    strWhere = "AND IIF(TRANSACTION_TYPE = 1,PAYBACK_DATE,EVENT_DATE) BETWEEN  @duringStart AND  @duringEnd ";
                    parameters.Add(new SqlParameter("duringStart", duringStart));
                    parameters.Add(new SqlParameter("duringEnd", duringEnd));
                    logger.Debug("duringStart,duringEnd=" + duringStart + "," + duringEnd);
                }
                sql = sql.Replace("$WhereCond", strWhere);
                sql = sql + " ORDER BY IIF(TRANSACTION_TYPE = 1,PAYBACK_DATE,EVENT_DATE)";
                logger.Debug("sql=" + sql);
                lstItem = context.Database.SqlQuery<LoanTranactionFunction>(sql, parameters.ToArray()).ToList();
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
                    string sql = @"SELECT p.* , CONVERT(char(10), p.PAYMENT_DATE, 111) AS RECORDED_DATE, PARSENAME(Convert(varchar,Convert(money,ISNULL(p.AMOUNT_PAID, 0)),1),2) AS RECORDED_AMOUNT_PAYABLE, 
                            PARSENAME(Convert(varchar, Convert(money, ISNULL(it.AMOUNT, 0)), 1), 2) AS PAYBACK_AMOUNT, PARSENAME(Convert(varchar, Convert(money, ISNULL(p.AMOUNT_PAID, 0) - ISNULL(it.AMOUNT, 0)), 1), 2) AS RECORDED_AMOUNT_PAID  
                            FROM (
                            select o.PAYEE, CONVERT(datetime,o.PAYMENT_DATE, 111)PAYMENT_DATE, SUM(o.AMOUNT)AMOUNT_PAID 
                            from(SELECT CONVERT(varchar, PAYMENT_DATE, 111)PAYMENT_DATE, SUM(fei.AMOUNT)AMOUNT, PAYEE 
                            FROM FIN_EXPENSE_FORM fef LEFT JOIN FIN_EXPENSE_ITEM fei ON fef.EXP_FORM_ID = fei.EXP_FORM_ID 
                            WHERE fef.STATUS = 30 AND CONVERT(varchar, PAYMENT_DATE, 111) = @paymentDate 
                            GROUP BY CONVERT(varchar, PAYMENT_DATE, 111), PAYEE 
                            union 
                            SELECT CONVERT(char(10), PAYMENT_DATE, 111)PAYMENT_DATE, SUM(AMOUNT_PAID)AMOUNT_PAID, PAYEE 
                            FROM PLAN_ACCOUNT WHERE ISDEBIT = 'N' 
                            AND CONVERT(char(10), PAYMENT_DATE, 111) = @paymentDate 
                            GROUP BY PAYEE, CONVERT(char(10), PAYMENT_DATE, 111))o GROUP BY o.PAYEE, o.PAYMENT_DATE
                            )p LEFT JOIN(
                            SELECT SUM(t.AMOUNT)AMOUNT, l.BANK_NAME FROM FIN_LOAN_TRANACTION t LEFT JOIN FIN_BANK_LOAN l ON t.BL_ID = l.BL_ID  
                            WHERE ISNULL(l.IS_SUPPLIER, 'N') = 'Y' 
                            AND t.TRANSACTION_TYPE = 1 
                            OR ISNULL(l.IS_SUPPLIER, 'N') <> 'Y' 
                            AND t.REMARK NOT LIKE '%備償%' 
                            AND t.TRANSACTION_TYPE = -1 
                            AND CONVERT(char(10),IIF(TRANSACTION_TYPE = 1, PAYBACK_DATE, EVENT_DATE), 111) = @paymentDate 
                            GROUP BY l.BANK_NAME
                            )it ON it.BANK_NAME = p.PAYEE";
                    lstItem = context.Database.SqlQuery<PlanAccountFunction>(sql, new SqlParameter("paymentDate", paymentDate)).ToList();
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
        /// <summary>
        /// 取得預算數、實際數
        /// </summary>
        /// <param name="paymentDate"></param>
        /// <param name="duringStart"></param>
        /// <param name="duringEnd"></param>
        /// <returns></returns>
        public List<ExpenseFormFunction> getExpenseBudgetByDate(string paymentDate, string duringStart, string duringEnd)
        {
            List<ExpenseFormFunction> lstItem = new List<ExpenseFormFunction>();
            string sql = @"SELECT CONVERT(char(10),CONVERT(datetime, exp.PAID_DATE , 111), 111) RECORDED_DATE, SUM(exp.AMOUNT) AMOUNT
                FROM (
                	SELECT (SELECT TOP 1 PROJECT_NAME FROM TND_PROJECT WHERE PROJECT_ID=v.PROJECT_ID ) as PROJECT_ID
                   ,(SELECT TOP 1 SUBJECT_NAME FROM FIN_SUBJECT WHERE FIN_SUBJECT_ID=v.SUBJECT_ID ) as SUBJECT_ID
                   ,PAID_DATE,AMOUNT,AMOUNT_REAL 
                   FROM vw_BudgetStatus v $WhereCond
                ) exp GROUP BY PAID_DATE
                ";
            using (var context = new topmepEntities())
            {
                //條件篩選
                if (null != paymentDate && paymentDate != "")
                {
                    string sql4Where = @"WHERE PAID_DATE = @paymentDate ";
                    sql = sql.Replace("$WhereCond", sql4Where);
                    logger.Debug("sql=" + sql + ",paymentdate=" + paymentDate);
                    lstItem = context.Database.SqlQuery<ExpenseFormFunction>(sql, new SqlParameter("paymentDate", paymentDate)).ToList();
                }
                else if (null != duringStart && duringStart != "" && null != duringEnd && duringEnd != "")
                {
                    string sql4Where = @"WHERE PAID_DATE BETWEEN @duringStart AND  @duringEnd";
                    sql = sql.Replace("$WhereCond", sql4Where);
                    logger.Debug("sql=" + sql + ",duringStart=" + duringStart + ",duringEnd=" + duringEnd);
                    lstItem = context.Database.SqlQuery<ExpenseFormFunction>(sql, new SqlParameter("duringStart", duringStart), new SqlParameter("duringEnd", duringEnd)).ToList();
                }
                else
                {
                    string sql4Where = @"";
                    sql = sql.Replace("$WhereCond", sql4Where);
                    logger.Debug("sql=" + sql);
                    lstItem = context.Database.SqlQuery<ExpenseFormFunction>(sql).ToList();
                }
            }
            return lstItem;
        }
        public List<Budget4CashFow> getBudgetStatus(string paymentDate)
        {
            List<Budget4CashFow> lstItem = new List<Budget4CashFow>();
            string sql = @"SELECT (SELECT TOP 1 PROJECT_NAME FROM TND_PROJECT WHERE PROJECT_ID=v.PROJECT_ID ) as PROJECT_ID
                        ,(SELECT TOP 1 SUBJECT_NAME FROM FIN_SUBJECT WHERE FIN_SUBJECT_ID=v.SUBJECT_ID ) as SUBJECT_ID
                        ,PAID_DATE,AMOUNT,AMOUNT_REAL 
                        FROM vw_BudgetStatus v WHERE  PAID_DATE=@paymentdate";
            using (var context = new topmepEntities())
            {
                //條件篩選
                if (null != paymentDate && paymentDate != "")
                {
                    logger.Debug("sql=" + sql + ",paymentdate=" + paymentDate);
                    lstItem = context.Database.SqlQuery<Budget4CashFow>(sql, new SqlParameter("paymentdate", paymentDate)).ToList();
                }
                return lstItem;
            }
        }
    }
    /// <summary>
    /// 支付與估驗服務
    /// </summary>
    class Service4Payment: PurchaseFormService
    {

    }
}

using log4net;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using topmeperp.Models;
using System.Data.Entity.Migrations;

namespace topmeperp.Service
{
    //銀行帳戶服務層
    public class ContextService4BankInfo : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        //封裝供前端頁面調用
        public TndProjectModels tndProjectModels = new TndProjectModels();
        public ContextService4BankInfo()
        {

        }
        //增加銀行帳戶
        public int addBankInfo(FIN_BANK_ACCOUNT account)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.FIN_BANK_ACCOUNT.Add(account);
                    i = context.SaveChanges();
                    logger.Info("new bank account record=" + account.BANK_ACCOUNT_ID + ",initial amount=" + account.CUR_AMOUNT + ",curDate=" + account.CUR_DATE);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":StackTrace=" + ex.StackTrace);
                }
            }
            logger.Info("add bankAccount count =" + i);
            return i;
        }
        //取得銀行帳戶
        public List<FIN_BANK_ACCOUNT> getAllBankAccount()
        {
            List<FIN_BANK_ACCOUNT> lstBankAccount = null;
            using (var context = new topmepEntities())
            {
                try
                {
                    lstBankAccount = context.FIN_BANK_ACCOUNT.ToList();
                    logger.Info("new bank account records=" + lstBankAccount.Count);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":StackTrace=" + ex.StackTrace);
                }
            }
            return lstBankAccount;
        }
        //更新帳戶現額資料
        public void updateBankAccount(List<FIN_BANK_ACCOUNT> lstBankAccount)
        {
            using (var context = new topmepEntities())
            {
                string sql = "UPDATE FIN_BANK_ACCOUNT SET CUR_AMOUNT=@CurAmt,CUR_DATE=@CurDate,MODIFY_ID=@ModifyId,MODIFY_DATE=@ModifyDate WHERE BANK_ACCOUNT_ID=@BankAccountId";
                foreach (FIN_BANK_ACCOUNT account in lstBankAccount)
                {
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("BankAccountId", account.BANK_ACCOUNT_ID));
                    parameters.Add(new SqlParameter("CurAmt", account.CUR_AMOUNT));
                    parameters.Add(new SqlParameter("CurDate", account.CUR_DATE));
                    parameters.Add(new SqlParameter("ModifyId", account.MODIFY_ID));
                    parameters.Add(new SqlParameter("ModifyDate", account.MODIFY_DATE));
                    logger.Info("update bank account record=" + account.BANK_ACCOUNT_ID + ",initial amount=" + account.CUR_AMOUNT + ",curDate=" + account.CUR_DATE);
                    context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                }
            }
        }
        //取得貸款帳戶資料
        public List<BankLoanInfoExt> getAllBankLoan()
        {
            List<BankLoanInfoExt> lstBankLoan = null;
            using (var context = new topmepEntities())
            {
                try
                {
                    string sql = @"
                    SELECT B.* , P.PROJECT_NAME, ISNULL(ROUND(V.vaRatio * 100, 0),0) AS vaRatio  , 
                    (SELECT ISNULL(IIF(QUOTA_RECYCLABLE = 'Y', SUM(TRANSACTION_TYPE*AMOUNT),(SUM(IIF(TRANSACTION_TYPE = 1,0,-1)*AMOUNT))), 0) 
                    FROM FIN_LOAN_TRANACTION T WHERE T.BL_ID = B.BL_ID) SumTransactionAmount 
                    FROM FIN_BANK_LOAN B LEFT JOIN TND_PROJECT P ON B.PROJECT_ID = P.PROJECT_ID LEFT JOIN
                    (SELECT va.PROJECT_ID AS PROJECT_ID, ISNULL(va.VALUATION_AMOUNT, 0) / ISNULL(pi.contractAtm, 1) AS vaRatio 
                    FROM (SELECT PROJECT_ID, SUM(VALUATION_AMOUNT) AS VALUATION_AMOUNT 
                    FROM PLAN_VALUATION_FORM GROUP BY PROJECT_ID)va LEFT JOIN
                    (SELECT PROJECT_ID, SUM(ITEM_UNIT_PRICE * ITEM_QUANTITY) AS contractAtm 
                    FROM PLAN_ITEM GROUP BY PROJECT_ID)pi ON va.PROJECT_ID = pi.PROJECT_ID)V ON B.PROJECT_ID = V.PROJECT_ID WHERE ISNULL(IS_SUPPLIER, 'N') <> 'Y'
                    ";
                    logger.Debug("sql" + sql);
                    lstBankLoan = context.Database.SqlQuery<BankLoanInfoExt>(sql).ToList();
                    logger.Info("new bank loan records=" + lstBankLoan.Count);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":StackTrace=" + ex.StackTrace);
                }
            }
            return lstBankLoan;
        }
        
        //取得專案執行階段之專案
        public void getAllPlan()
        {
            List<TND_PROJECT> lstPlans = null;
            using (var context = new topmepEntities())
            {
                try
                {
                    lstPlans = context.TND_PROJECT.SqlQuery("SELECT * FROM TND_PROJECT WHERE STATUS = '專案執行' ").ToList();
                    logger.Debug("get records=" + lstPlans.Count);
                    //將專案執行階段所有專案封裝供前端頁面調用
                    tndProjectModels.planList = lstPlans;
                }
                catch (Exception e)
                {
                    logger.Error("fail:" + e.StackTrace);
                }
            }
        }
        //取得貸款帳戶交易資料
        public BankLoanInfo getBankLoan(string bl_id, string supplier)
        {
            BankLoanInfo item = new BankLoanInfo();
            using (var context = new topmepEntities())
            {
                try
                {
                    logger.Info("get bank transaction BL_ID=" + bl_id);
                    item.LoanInfo = context.FIN_BANK_LOAN.Find(long.Parse(bl_id));
                    long blid = long.Parse(bl_id);
                    item.LoanTransaction = context.FIN_LOAN_TRANACTION.Where(b => b.BL_ID == blid).ToList();
                    string sql = "";
                    //取得期數與匯總金額
                    if (supplier == "Y")
                    {
                        sql = "SELECT ISNULL(CUR_PERIOD, 0)CUR_PERIOD , ISNULL(AMOUNT, 0)AMOUNT, (SELECT ISNULL(SUM(AMOUNT), 0) FROM FIN_BANK_LOAN f " +
                            "LEFT JOIN FIN_LOAN_TRANACTION t ON f.BL_ID = t.BL_ID WHERE t.TRANSACTION_TYPE = 1 AND f.BL_ID =@BL_ID) AS paybackAmt, " +
                            "(SELECT ISNULL(SUM(AMOUNT), 0) FROM FIN_BANK_LOAN f LEFT JOIN FIN_LOAN_TRANACTION t ON f.BL_ID = t.BL_ID " +
                            "WHERE t.TRANSACTION_TYPE = -1 AND f.BL_ID =@BL_ID) AS eventAmt FROM FIN_BANK_LOAN f " +
                            "LEFT JOIN (SELECT BL_ID, MAX(ISNULL(PERIOD, 0)) CUR_PERIOD, SUM(TRANSACTION_TYPE * AMOUNT) AMOUNT " +
                            "FROM FIN_LOAN_TRANACTION GROUP BY BL_ID)flt ON flt.BL_ID = f.BL_ID WHERE f.BL_ID =@BL_ID ";
                    }
                    else
                    {
                        sql = "SELECT ISNULL(CUR_PERIOD, 0)CUR_PERIOD , ISNULL(AMOUNT, 0)AMOUNT, ROUND(ISNULL(fbl.QUOTA * (1-IIF(fbl.vaRatio >= fbl.CUM_AR_RATIO , 1, fbl.QUOTA_AVAILABLE_RATIO/100)), 0), 0) AS SURPLUS_AMOUNT, " +
                            "(SELECT ISNULL(SUM(AMOUNT), 0) FROM FIN_BANK_LOAN f LEFT JOIN FIN_LOAN_TRANACTION t ON f.BL_ID = t.BL_ID WHERE t.TRANSACTION_TYPE = 1 AND f.BL_ID =@BL_ID) AS paybackAmt, " +
                            "(SELECT ISNULL(SUM(AMOUNT), 0) FROM FIN_BANK_LOAN f LEFT JOIN FIN_LOAN_TRANACTION t ON f.BL_ID = t.BL_ID WHERE t.TRANSACTION_TYPE = -1 AND f.BL_ID =@BL_ID) AS eventAmt " +
                            "FROM (SELECT BL_ID, QUOTA, CUM_AR_RATIO, QUOTA_AVAILABLE_RATIO, ISNULL(VALUATION_AMOUNT,0) / ISNULL(contractAtm, 1) * 100 AS vaRatio FROM FIN_BANK_LOAN f LEFT JOIN " +
                            "(SELECT PROJECT_ID, SUM(ITEM_UNIT_PRICE * ITEM_QUANTITY) AS contractAtm FROM PLAN_ITEM GROUP BY PROJECT_ID)pi ON f.PROJECT_ID = pi.PROJECT_ID " +
                            "LEFT JOIN (SELECT PROJECT_ID, SUM(VALUATION_AMOUNT)AS VALUATION_AMOUNT FROM PLAN_VALUATION_FORM GROUP BY PROJECT_ID)v ON f.PROJECT_ID = v.PROJECT_ID " +
                            "WHERE BL_ID=@BL_ID)fbl LEFT JOIN (SELECT lt.BL_ID, MAX(ISNULL(PERIOD,0)) CUR_PERIOD, IIF(QUOTA_RECYCLABLE = 'Y',SUM(TRANSACTION_TYPE*AMOUNT), SUM(IIF(TRANSACTION_TYPE = 1,0,-1)*AMOUNT)) AMOUNT " +
                            "from FIN_LOAN_TRANACTION lt LEFT JOIN FIN_BANK_LOAN bl ON lt.BL_ID = bl.BL_ID GROUP BY lt.BL_ID, bl.QUOTA_RECYCLABLE)flt ON flt.BL_ID = fbl.BL_ID ";
                    }
                    logger.Debug("sql=" + sql);
                    Dictionary<string, object> para = new Dictionary<string, object>();
                    para.Add("BL_ID", blid);
                    DataSet ds = ExecuteStoreQuery(sql, System.Data.CommandType.Text, para);
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        item.CurPeriod = long.Parse(ds.Tables[0].Rows[0]["CUR_PERIOD"].ToString());
                        item.SumTransactionAmount = (decimal)ds.Tables[0].Rows[0]["AMOUNT"];
                        item.paybackAmt = (decimal)ds.Tables[0].Rows[0]["paybackAmt"];
                        item.eventAmt = (decimal)ds.Tables[0].Rows[0]["eventAmt"];
                        if (supplier != "Y")
                        {
                            item.SurplusQuota = (decimal)ds.Tables[0].Rows[0]["SURPLUS_AMOUNT"];
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":StackTrace=" + ex.StackTrace);
                }
            }
            return item;
        }
        //新增貸款帳戶資料
        public int addBankLoan(FIN_BANK_LOAN bankloan)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.FIN_BANK_LOAN.AddOrUpdate(bankloan);
                    i = context.SaveChanges();
                    logger.Info("new bank loan record=" + bankloan.ACCOUNT_NAME);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":StackTrace=" + ex.StackTrace);
                }
            }
            logger.Info("add bankloan count =" + i);
            return i;
        }
        /// <summary>
        /// 增加借款還款紀錄
        /// </summary>
        /// <param name="loanTransaction"></param>
        /// <returns></returns>
        public int addBankLoanTransaction(List<FIN_LOAN_TRANACTION> loanTransaction)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.FIN_LOAN_TRANACTION.AddRange(loanTransaction);
                    i = context.SaveChanges();
                    logger.Info("new bank loan transaction record=" + loanTransaction.Count);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":StackTrace=" + ex.StackTrace);
                }
            }
            return i;
        }
        //取得廠商借支資料
        public List<BankLoanInfoExt> getAllSupplierLoan()
        {
            List<BankLoanInfoExt> lstSupplierLoan = null;
            using (var context = new topmepEntities())
            {
                try
                {
                    string sql = "SELECT B.*, P.PROJECT_NAME, (SELECT ISNULL(SUM(TRANSACTION_TYPE * AMOUNT), 0) FROM FIN_LOAN_TRANACTION T WHERE T.BL_ID = B.BL_ID) SumTransactionAmount, " +
                        "isnull((SELECT ISNULL(SUM(AMOUNT), 0) FROM FIN_LOAN_TRANACTION f WHERE f.TRANSACTION_TYPE = 1 AND f.BL_ID = B.BL_ID GROUP BY f.BL_ID), 0) AS paybackAmt, " +
                        "isnull((SELECT ISNULL(SUM(AMOUNT), 0) FROM FIN_LOAN_TRANACTION f WHERE f.TRANSACTION_TYPE = -1 AND f.BL_ID = B.BL_ID GROUP BY f.BL_ID), 0) AS eventAmt " +
                        "FROM FIN_BANK_LOAN B LEFT JOIN TND_PROJECT P ON B.PROJECT_ID = P.PROJECT_ID WHERE ISNULL(B.IS_SUPPLIER, 'N') = 'Y'";

                    lstSupplierLoan = context.Database.SqlQuery<BankLoanInfoExt>(sql).ToList();
                    logger.Info("new supplier loan records=" + lstSupplierLoan.Count);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":StackTrace=" + ex.StackTrace);
                }
            }
            return lstSupplierLoan;
        }

        //取得銀行貸款單一交易資料
        public LoanTranactionFunction getLoanTransactionItem(string itemid)
        {
            logger.Debug("get bank loan transaction item by id=" + itemid);
            LoanTranactionFunction item = null;
            using (var context = new topmepEntities())
            {
                //條件篩選
                item = context.Database.SqlQuery<LoanTranactionFunction>("SELECT *, PARSENAME(Convert(varchar,Convert(money,ISNULL(AMOUNT, 0)),1),2) AS RECORDED_AMOUNT, " +
                    "CONVERT(char(10), EVENT_DATE, 111) AS RECORDED_EVENT_DATE,  CONVERT(char(10), PAYBACK_DATE, 111) AS RECORDED_PAYBACK_DATE, " +
                    "CONVERT(char(10), CREATE_DATE, 111) AS RECORDED_CREATE_DATE FROM FIN_LOAN_TRANACTION WHERE TID=@itemid",
                new SqlParameter("itemid", itemid)).First();
            }
            return item;
        }
        public string message = "";
        public int updateLoanTransactionItem(FIN_LOAN_TRANACTION item)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.FIN_LOAN_TRANACTION.AddOrUpdate(item);
                    i = context.SaveChanges();
                }
                catch (Exception e)
                {
                    logger.Error("update loan transaction item fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }
            }
            return i;
        }
    }
}

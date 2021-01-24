using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace topmeperp.Models
{
    public class BankLoanInfo
    {
        public FIN_BANK_LOAN LoanInfo { get; set; }
        public List<FIN_LOAN_TRANACTION> LoanTransaction { get; set; }
        public long CurPeriod { get; set; }
        public decimal SumTransactionAmount { get; set; }
        public decimal SurplusQuota { get; set; }
        public decimal paybackAmt { get; set; }
        public decimal eventAmt { get; set; }

    }
    public class BankLoanInfoExt: FIN_BANK_LOAN
    {
        public decimal SumTransactionAmount { get; set; }
        public decimal vaRatio { get; set; }
        public decimal paybackAmt { get; set; }
        public decimal eventAmt { get; set; }
        public string PROJECT_NAME { get; set; }
    }
    public class Budget4CashFow
    {
        public string PROJECT_ID { get; set; }
        public string SUBJECT_ID { get; set; }
        public DateTime PAID_DATE { get; set; }
        public Nullable<decimal> AMOUNT { get; set; }
        public Nullable<decimal> AMOUNT_REAL { get; set; }
    }
}
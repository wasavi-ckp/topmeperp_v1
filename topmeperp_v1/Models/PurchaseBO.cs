using System;
using System.Collections.Generic;

namespace topmeperp.Models
{
    //採發階段之發包標的
    public class PURCHASE_ORDER: PLAN_SUP_INQUIRY
    {
        public Nullable<decimal> BudgetAmount { get; set; }
        public Nullable<int> CountPO { get; set; }
        public Int64 NO { get; set; }
        public string Bargain { get; set; }
        public Nullable<decimal> paymentFrequency { get; set; }
        public string PAYMENT_TERMS { get; set; }
        public string ContractId { get; set; }
        public string Supplier { get; set; }
        public string MATERIAL_BRAND { get; set; }
        public string CONTRACT_PRODUCTION { get; set; }
        public string DELIVERY_DATE { get; set; }
        public string ConRemark { get; set; }

    }
    public class BUDGET_SUMMANY 
    {
        public Nullable<decimal> Material_Budget { get; set; }
        public Nullable<decimal> Wage_Budget { get; set; }
    }


}
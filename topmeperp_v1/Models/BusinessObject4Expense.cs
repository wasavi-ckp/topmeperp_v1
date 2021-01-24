using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace topmeperp.Models
{
    //
    public class ExpenseTask
    {
        //任務現況
        public ExpenseFlowTask task { get; set; }
        //對應到表單的請求資料
        public WF_PROCESS_REQUEST ProcessRequest { get; set; }
        //表單簽核所需的步驟
        public List<WF_PORCESS_TASK> ProcessTask { get; set; }
        //表單內容
        public OperatingExpenseModel FormData { get; set; }
        public ContractModels EstData { get; set; }
        //申請者部門資料
        public RequestUserDeptInfo DeptInfo { get; set; }
    }

    //提供費用申請流程清單所需物件
    public class ExpenseFlowTask: WF_PORCESS_TASK
    {
        //FIN_EXPENSE_FORM
        public string EXP_FORM_ID { get; set; }
        public string EST_FORM_ID { get; set; }
        public string PROJECT_ID { get; set; }
        public Int32 OCCURRED_YEAR { get; set; }
        public Int32 OCCURRED_MONTH { get; set; }
        ///廠商編號
        public string PAYEE { get; set; }
        //廠商名稱
        public string SUPPLIER_NAME { get; set; }
        public DateTime? PAYMENT_DATE { get; set; }
        public string REQ_DESC { get; set; }
        public string REJECT_DESC { get; set; }

        //WF_PROCESS_REQUEST
        public string REQ_USER_ID { get; set; }
        public Int64 CURENT_STATE { get; set; }
        public Int64 PID { get; set; } //可關聯至WF_PROCESS
        // FROM URL
        public string FORM_URL { get; set; }
        //DE     
        public string DEPT_CODE { get; set; } //申請者部門
        public string MANAGER { get; set; }//申請部門主管
        public decimal PAID_AMOUNT { get; set; }
        public string FORM_NAME { get; set; }
        public string PROJECT_NAME { get; set; }
    }
    //申請者部門資料
    public class RequestUserDeptInfo
    {
        public string REQ_USER_ID { get; set; }
        public string REQ_USER_NAME { get; set; }
        public string DEPT_CODE { get; set; }
        public string DEPT_NAME { get; set; }
        public string MANAGER { get; set; }
        public string MANAGER_NAME { get; set; }
    }
    public class EstmTask:ExpenseTask {
        public new ContractModels FormData { get; set; }
    }
}
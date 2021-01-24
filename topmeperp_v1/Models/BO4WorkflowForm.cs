using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace topmeperp.Models
{
    public class CostChangeFormTask:ExpenseTask
    {
        //表單內容
        public new PLAN_COSTCHANGE_FORM FormData { get; set; }
        public List<PLAN_COSTCHANGE_ITEM> lstItem { get; set; }
    }
    //提供費用申請流程清單所需物件
    public class CostChangeTask : WF_PORCESS_TASK
    {
        //PLAN_COSTCHANGE_FORM
        public string FORM_ID { get; set; }
        public string PROJECT_ID { get; set; }
        public string REJECT_DESC { get; set; }
        public string REMARK_ITEM { get; set; }
        public string REMARK_QTY { get; set; }
        public string REMARK_PRICE { get; set; }
        public string REMARK_OTHER { get; set; }
        //WF_PROCESS_REQUEST
        public string REQ_USER_ID { get; set; }
        public Int64 CURENT_STATE { get; set; }
        public Int64 PID { get; set; } //可關聯至WF_PROCESS
        // FROM URL
        public string FORM_URL { get; set; }
        //DE     
        public string DEPT_CODE { get; set; } //申請者部門
        public string MANAGER { get; set; }//申請部門主管
    }
}
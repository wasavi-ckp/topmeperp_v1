using log4net;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    public class EMailService
    {
        static ILog log = log4net.LogManager.GetLogger(typeof(EMailService));
        string smtp_ip = "";
        int smtp_port = 587;
        string smpt_id = "";
        string smtp_pwd = "";
        string system_url = "";
        public EMailService()
        {
            smtp_ip = ConfigurationManager.AppSettings["smtp_ip"];
            smtp_port = int.Parse(ConfigurationManager.AppSettings["smtp_port"].ToString());
            smpt_id = ConfigurationManager.AppSettings["smtp_id"];
            smtp_pwd = ConfigurationManager.AppSettings["smtp_pwd"];
            system_url = ConfigurationManager.AppSettings["system_url"];
            log.Info("smtp_ip=" + smtp_ip + "smtp_port=" + smtp_port + "smpt_id=" + smpt_id);
        }
        public bool SendMailByGmail(string fromAddress, string sendername, string MailList, string bccMailList, string Subject, string Body, string filePath)
        {
            MailMessage msg = new MailMessage();
            //收件者，以逗號分隔不同收件者 ex "test@gmail.com,test2@gmail.com"
            if (null != MailList && "" != MailList)
            {
                msg.To.Add(MailList);
            }
            //密件副本清單
            if (null != bccMailList && "" != bccMailList)
            {
                msg.Bcc.Add(bccMailList);
            }

            msg.From = new MailAddress(fromAddress, sendername, System.Text.Encoding.UTF8);
            //郵件標題 
            msg.Subject = Subject;
            //郵件標題編碼  
            msg.SubjectEncoding = System.Text.Encoding.UTF8;
            //郵件內容
            msg.Body = Body;
            msg.IsBodyHtml = true;
            //  mailservice.SendMailByGmail
            if (null != filePath)
            {
                Attachment attachment = new Attachment(filePath);
                msg.Attachments.Add(attachment);
            }
            msg.BodyEncoding = System.Text.Encoding.UTF8;//郵件內容編碼 
            msg.Priority = MailPriority.Normal;//郵件優先級 
                                               //建立 SmtpClient 物件 並設定 Gmail的smtp主機及Port 
            #region 其它 Host
            /*
             *  outlook.com smtp.live.com port:25
             *  yahoo smtp.mail.yahoo.com.tw port:465
            */
            #endregion
            SmtpClient MySmtp = new SmtpClient(smtp_ip, smtp_port);
            //設定你的帳號密碼
            MySmtp.Credentials = new System.Net.NetworkCredential(smpt_id, smtp_pwd);
            //Gmial 的 smtp 使用 SSL
            MySmtp.EnableSsl = true;
            log.Info("Email Send!!");
            try
            {
                MySmtp.Send(msg);
                return true;
            }
            catch (Exception ex)
            {
                log.Error(ex.StackTrace);
                return false;
            }
        }
        public void createPRMessage(SYS_USER u, PLAN_PURCHASE_REQUISITION pr)
        {
            //mail to 業管
            // System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            //  string userJson = objSerializer.Serialize(u);
            // log.Debug("sender :" + userJson);

            //定義Email 樣板
            StringBuilder strTemp = new StringBuilder("{1} 提出新的申購單");
            StringBuilder strTempBody = new StringBuilder("{1} 提出新的申購單{2}。 <br/> 系統發出，請勿回覆!!");

            SYS_MESSAGE m = new SYS_MESSAGE();
            m.FROM_ADDRESS = u.EMAIL;
            m.DISPLAY_NAME = u.USER_NAME;
            UserService s = new UserService();
            List<SYS_USER> lstTarget = s.getProjectUser(pr.PROJECT_ID, "業管");
            strTemp.Replace("{1}", pr.PROJECT_ID);

            m.MAIL_LIST = getMailList(lstTarget);
            if (m.MAIL_LIST == "")
            {
                log.Error(pr.PR_ID + " Have no mail list!!");
                return;
            }
            m.SUBJECT = strTemp.ToString();

            strTempBody.Replace("{1}", pr.PROJECT_ID);
            strTempBody.Replace("{2}", "申購單號為:<a href='" + system_url + "'>" + pr.PR_ID + "</a>");
            m.MSG_BODY = strTempBody.ToString();
            log.Debug("email body:" + m.MSG_BODY);
            m.CREATE_ID = u.USER_ID;
            m.CREATE_DATE = DateTime.Now;
            //決定發送時間
            m.SEND_TIME = DateTime.Now;
            int i = 0;
            using (var context = new topmepEntities())
            {

               context.SYS_MESSAGE.Add(m);
                i = context.SaveChanges();
            }
        }
        public void createPOMessage(SYS_USER u, PLAN_PURCHASE_REQUISITION pr)
        {
            //mail to 業管
            // System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            //  string userJson = objSerializer.Serialize(u);
            // log.Debug("sender :" + userJson);

            //定義Email 樣板
            StringBuilder strTemp = new StringBuilder("{1} 申購單已完成採購!!");
            StringBuilder strTempBody = new StringBuilder("{1} 申購單已完成採購{2}。 <br/> 系統發出，請勿回覆!!");
            //
            SYS_MESSAGE m = new SYS_MESSAGE();
            m.FROM_ADDRESS = u.EMAIL;
            m.DISPLAY_NAME = u.USER_NAME;
            UserService s = new UserService();
            //TODO:加入通知申請人
            List<SYS_USER> lstTarget = s.getProjectUser(pr.PROJECT_ID, "工地主任");
            strTemp.Replace("{1}", pr.PROJECT_ID);

            m.MAIL_LIST = getMailList(lstTarget);
            if (m.MAIL_LIST == "")
            {
                log.Error(pr.PR_ID + " Have no mail list!!");
                return;
            }
            m.SUBJECT = strTemp.ToString();

            strTempBody.Replace("{1}", pr.PROJECT_ID);
            strTempBody.Replace("{2}", "採購單號為:<a href='" + system_url + "'>" + pr.PR_ID + "</a>");
            m.MSG_BODY = strTempBody.ToString();
            log.Debug("email body:" + m.MSG_BODY);
            m.CREATE_ID = u.USER_ID;
            m.CREATE_DATE = DateTime.Now;
            //決定發送時間
            m.SEND_TIME = DateTime.Now;
            int i = 0;
            using (var context = new topmepEntities())
            {

                context.SYS_MESSAGE.Add(m);
                i = context.SaveChanges();
            }
        }
        private string getMailList(List<SYS_USER> lstTarget)
        {
            string MailLis = "";
            foreach (SYS_USER targetUser in lstTarget)
            {
                if (MailLis == "")
                {
                    MailLis = targetUser.EMAIL;
                }
                else
                {
                    MailLis = MailLis + "," + targetUser.EMAIL;
                }
            }
            return MailLis;
        } 
    }

}
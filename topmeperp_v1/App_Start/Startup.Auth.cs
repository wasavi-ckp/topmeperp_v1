using Microsoft.AspNet.Identity;
using Microsoft.Owin;
using Microsoft.Owin.Security.Cookies;
using Owin;
using log4net;
using System.Diagnostics;
using System.Reflection;
using Hangfire;
using topmeperp.Models;
using Hangfire.SqlServer;
using System;
using topmeperp.Schedule;

namespace topmeperp
{
    public partial class Startup
    {
        ILog log = log4net.LogManager.GetLogger(typeof(Startup));
        // 如需設定驗證的詳細資訊，請瀏覽 http://go.microsoft.com/fwlink/?LinkId=301864
        public void ConfigureAuth(IAppBuilder app)
        {
            // 配置Middleware 組件
            log.Info("StartupConfigureAuth");
            app.UseCookieAuthentication(new CookieAuthenticationOptions
            {
                AuthenticationType = DefaultAuthenticationTypes.ApplicationCookie,
                LoginPath = new PathString("/Index/Login"),
                CookieSecure = CookieSecureOption.Never,
            });
            //加入HangFire 排程作業
            var context = new topmepEntities();
            //使用ms sql 紀錄任務
            GlobalConfiguration.Configuration.UseSqlServerStorage(context.Database.Connection.ConnectionString,
                new SqlServerStorageOptions
                {
                    // if it is set to 1 minutes, each worker will run a keep-alive query each minute when processing a job
                    QueuePollInterval = TimeSpan.FromMinutes(1)
                }

                );
            // 啟用HanfireServer
            app.UseHangfireServer(new BackgroundJobServerOptions { WorkerCount = 1 });
            RecurringJob.AddOrUpdate<BackgroundService>("BackgroundService", x => x.SendMailSchedule(),
                Cron.MinuteInterval(1),
                TimeZoneInfo.FindSystemTimeZoneById("Taipei Standard Time"));
            // 啟用Hangfire的Dashboard
            app.UseHangfireDashboard();

        }
    }
    public class AppInfo
    {
        public static string Version = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion.ToString();
    }
}
using System.Collections;
using java.util;
using net.sf.mpxj;
using System;
using System.Collections.Generic;
using topmeperp.Models;
using log4net;
using System.Data.SqlClient;

namespace topmeperp.Service
{
    public class OfficeProjectService : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string project_id;
        List<PLAN_TASK> lstTask = null;
        public void convertProject(string projectId, string filepath)
        {
            project_id = projectId;
            net.sf.mpxj.mpp.MPPReader reader = new net.sf.mpxj.mpp.MPPReader();
            ProjectFile projectObj = reader.read(filepath);
            int i = 1;
            lstTask = new List<PLAN_TASK>();
            foreach (net.sf.mpxj.Task task in ToEnumerable(projectObj.AllTasks))
            {
                PLAN_TASK pt = new PLAN_TASK();
                pt.PROJECT_ID = projectId;

                pt.PRJ_UID = task.UniqueID.intValue();
                pt.PRJ_ID = task.ID.intValue();
                pt.TASK_NAME = task.Name;

                DateTime dtStart = new DateTime();
                DateTime dtFinish = new DateTime();
                //ToString("yyyyMMddHHmmss")
                if (null != task.Start)
                {
                    dtStart = new DateTime((task.Start.getYear() + 1900), task.Start.getMonth() + 1, task.Start.getDate());
                    logger.Debug("start date Year =" + (task.Start.getYear() + 1900) + ",Month=" + (task.Start.getMonth() + 1) + ",Date=" + task.Start.getDate());
                    pt.START_DATE = dtStart;

                    dtFinish = new DateTime((task.Finish.getYear() + 1900), task.Finish.getMonth() + 1, task.Finish.getDate());
                    pt.FINISH_DATE = dtFinish;
                    logger.Debug("start date Year =" + (task.Finish.getYear() + 1900) + ",Month=" + (task.Finish.getMonth() + 1) + ",Date=" + task.Finish.getDate());

                    pt.DURATION = task.Duration.toString();
                }
                logger.Debug("DURATION=" + task.Duration + ",Task: " + i + "=" + task.Name + ",StartDate=" + dtStart.ToString("yyyy/MM/dd") + ",EndDate=" + dtFinish.ToString("yyyy/MM/dd") + " ID=" + task.ID + " Unique ID=" + task.UniqueID);
                if (null != task.ParentTask)
                {
                    pt.PARENT_UID = task.ParentTask.UniqueID.intValue();
                }
                logger.Info("Parent UID=" + pt.PARENT_UID + ",TASK_UID=" + pt.PRJ_UID + ",Task_id=" + pt.TASK_ID + ",TASK_NAME=" + pt.TASK_NAME);
                lstTask.Add(pt);
                i++;
            }
            logger.Info("Get all task count:" + lstTask.Count);
        }
        public void import2Table()
        {
            if (null != lstTask)
            {
                using (var context = new topmepEntities())
                {
                    //1.清除所有任務
                    string sql = "DELETE FROM PLAN_TASK WHERE PROJECT_ID=@projectid";
                    int i = context.Database.ExecuteSqlCommand(sql, new SqlParameter("projectid", project_id));
                    logger.Debug("Remove Exist Task for projectid=" + project_id);
                    //2.匯入任務
                    foreach (PLAN_TASK pt in lstTask)
                    {
                        if (pt.TASK_NAME != null)
                        {
                            context.PLAN_TASK.Add(pt);
                        }else
                        {
                            logger.Warn("task name is null:" + pt.PRJ_UID + ",id=" + pt.PRJ_ID);
                        }
                    }
                    i = context.SaveChanges();
                    logger.Info("import task count=" + i);
                }
            }
        }
        private static EnumerableCollection ToEnumerable(Collection javaCollection)
        {
            return new EnumerableCollection(javaCollection);
        }
    }

    public class EnumerableCollection
    {
        public EnumerableCollection(Collection collection)
        {
            m_collection = collection;
        }

        public IEnumerator GetEnumerator()
        {
            return new Enumerator(m_collection);
        }

        private Collection m_collection;
    }

    public class Enumerator : IEnumerator
    {
        public Enumerator(Collection collection)
        {
            m_collection = collection;
            m_iterator = m_collection.iterator();
        }

        public object Current
        {
            get
            {
                return m_iterator.next();
            }
        }

        public bool MoveNext()
        {
            return m_iterator.hasNext();
        }

        public void Reset()
        {
            m_iterator = m_collection.iterator();
        }

        private Collection m_collection;
        private Iterator m_iterator;
    }
}
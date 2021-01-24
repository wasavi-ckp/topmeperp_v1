using log4net;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Core.Objects;
using System.Data.Entity.Migrations;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    #region 部門資料管理管理區塊
    /*
     *部門資料管理管理 */

    public class DepartmentManage : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public ENT_DEPARTMENT department = null;
        public List<ENT_DEPARTMENT> Org_department = null;

        //取得部門資料
        public ENT_DEPARTMENT getDepartmentById(string depid)
        {
            using (var context = new topmepEntities())
            {
                department = context.ENT_DEPARTMENT.SqlQuery("select dep.* from ENT_DEPARTMENT dep "
                    + "where dep.DEP_ID = @depid "
                   , new SqlParameter("depid", depid)).First();
            }
            return department;
        }
        //新增部門資料
        public long addDepartment(ENT_DEPARTMENT dep)
        {
            logger.Info("create new Department ");
            using (var context = new topmepEntities())
            {
                //2.取得供應商編號
                context.ENT_DEPARTMENT.AddOrUpdate(dep);
                int i = context.SaveChanges();
                logger.Debug("Add dep=" + i);
            }
            return dep.DEP_ID;
        }
        /// <summary>
        /// 刪除部門資料
        /// </summary>
        /// <param name="depId"></param>
        public void delDepartment(long depId)
        {
            using (var context = new topmepEntities())
            {
                //2.取得
                ENT_DEPARTMENT d= context.ENT_DEPARTMENT.Find(depId);
                logger.Info("del Department =" + d.DEPT_NAME +"," +d.DEP_ID);
                context.ENT_DEPARTMENT.Remove(d);
                int i = context.SaveChanges();
                logger.Debug("remove dep=" + i);
            }
        }
        /// <summary>
        /// 取得所有部門資料(CTE)
        /// </summary>
        /// <param name="depid"></param>
        /// <returns></returns>
        public List<ENT_DEPARTMENT> getDepartmentOrg(int depid)
        {
            string sql = @"WITH Dep_CTE (DEP_ID,DEPT_CODE,DEPT_NAME,PARENT_ID,DEP_LEVEL) AS
                (    
                    --頂層
                    select DEP_ID,DEPT_CODE,DEPT_NAME,PARENT_ID,0 AS DT_LEVEL
                    from ENT_DEPARTMENT RD
                    where RD.DEP_ID=@depid
                    UNION ALL    
                    --成員
                    select D.DEP_ID,D.DEPT_CODE,D.DEPT_NAME,D.PARENT_ID,DC.DEP_LEVEL+1
                    from ENT_DEPARTMENT D INNER JOIN Dep_CTE DC on D.PARENT_ID=DC.DEP_ID
                )
                select * from Dep_CTE";
            var parameters = new Dictionary<string, Object>();
            //設定專案名編號資料
            parameters.Add("depid", depid);
            DataSet ds = ExecuteStoreQuery(sql, CommandType.Text, parameters);
            Org_department = new List<ENT_DEPARTMENT>();
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                ENT_DEPARTMENT dept = new ENT_DEPARTMENT();
                dept.DEP_ID = Convert.ToInt64(dr["DEP_ID"]);
                dept.DEPT_CODE = Convert.ToString(dr["DEPT_CODE"]);
                dept.DEPT_NAME = Convert.ToString(dr["DEPT_NAME"]);
                //dept.DESC = Convert.ToString(dr["DESC"]);
                dept.PARENT_ID = Convert.ToInt64(dr["PARENT_ID"]);
                Org_department.Add(dept);
            }
            return Org_department;
        }

        /// <summary>
        /// 取得資料
        /// </summary>
        /// <param name="depid"></param>
        /// <returns></returns>
        public string getDepartment4Tree(int depid)
        {
            string sql = "SELECT * FROM ENT_DEPARTMENT ORDER BY PARENT_ID";
            List<ENT_DEPARTMENT> rtDept = new List<ENT_DEPARTMENT>();
            using (var context = new topmepEntities())
            {
                rtDept = context.ENT_DEPARTMENT.SqlQuery(sql).ToList();
                logger.Debug("row count=" + rtDept.Count);
            }
            // Dictionary<int, PROJECT_TASK_TREE_NODE> dicTree = new Dictionary<int, PROJECT_TASK_TREE_NODE>();
            //PROJECT_TASK_TREE_NODE rootnode = new PROJECT_TASK_TREE_NODE();
            Dictionary<int, DEPARTMENT_TREE4SHOW> dicTree = new Dictionary<int, DEPARTMENT_TREE4SHOW>();
            DEPARTMENT_TREE4SHOW rootnode = new DEPARTMENT_TREE4SHOW();
            foreach (ENT_DEPARTMENT t in rtDept)
            {
                //將跟節點置入Directory 內
                if (t.PARENT_ID == 0 || dicTree.Count == 0)
                {
                    rootnode.href = t.DEP_ID.ToString();
                    rootnode.text = t.DEP_ID + "-" + t.DEPT_CODE + "-" + t.DEPT_NAME;
                    dicTree.Add(Convert.ToInt32(t.DEP_ID), rootnode);
                    logger.Info("add root node :" + t.DEP_ID);
                }
                else
                {
                    //將Dic 內的節點翻出，加入子節點
                    DEPARTMENT_TREE4SHOW parentnode = (DEPARTMENT_TREE4SHOW)dicTree[Convert.ToInt32(t.PARENT_ID)];
                    DEPARTMENT_TREE4SHOW node = new DEPARTMENT_TREE4SHOW();

                    node.href = t.DEP_ID.ToString();
                    node.text = t.DEP_ID + "-" + t.DEPT_CODE + "-" + t.DEPT_NAME;
                    node.tags.Add("主管:" + t.MANAGER);
                    parentnode.addChild(node);
                    //將結點資料記錄至dic 內
                    dicTree.Add(Convert.ToInt32(t.DEP_ID), node);
                    logger.Info("add  node :" + t.DEP_ID + ",parent=" + t.PARENT_ID);
                }
            }
            string output = JsonConvert.SerializeObject(rootnode);
            logger.Info("Jason:" + output);
            return output;
        }
    }
    #endregion
}

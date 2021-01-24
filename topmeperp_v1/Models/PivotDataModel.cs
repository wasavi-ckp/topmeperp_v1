using log4net;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace topmeperp.Models
{
    #region pivot sample 1
    public class PivotDataModel
    {


    }
    public class PivotRow<TypeId, TypeAttr, TypeValue>
    {
        public TypeId ObjectId { get; set; }
        public IEnumerable<TypeAttr> Attributes { get; set; }
        public IEnumerable<TypeValue> Values { get; set; }


        public static DataTable GetPivotTable(List<TypeAttr> attributeNames, List<PivotRow<TypeId, TypeAttr, TypeValue>> source)
        {
            DataTable dt = new DataTable();
            DataColumn dc = new DataColumn("ID", typeof(TypeId));
            dt.Columns.Add(dc);
            // Creat the new DataColumn for each attribute 
            attributeNames.ForEach(name =>
            {
                dc = new DataColumn(name.ToString(), typeof(TypeValue));
                dt.Columns.Add(dc);
            });
            // Insert the data into the Pivot table 
            foreach (PivotRow<TypeId, TypeAttr, TypeValue> row in source)
            {
                DataRow dr = dt.NewRow();
                dr["ID"] = row.ObjectId;
                List<TypeAttr> attributes = row.Attributes.ToList();
                List<TypeValue> values = row.Values.ToList();
                // Set the value basing the attribute names. 
                for (int i = 0; i < values.Count; i++)
                {
                    dr[attributes[i].ToString()] = values[i];
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
    #endregion
    #region pivot sample 2
    public class Pivot
    {
        private DataTable _SourceTable = new DataTable();
        private IEnumerable<DataRow> _Source = new List<DataRow>();
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public Pivot(DataTable SourceTable)
        {
            _SourceTable = SourceTable;
            _Source = SourceTable.Rows.Cast<DataRow>();
        }

        /// <summary>
        /// Pivots the DataTable based on provided RowField, DataField, Aggregate Function and ColumnFields.//
        /// </summary>
        /// <param name="rowField">The column name of the Source Table which you want to spread into rows</param>
        /// <param name="dataField">The column name of the Source Table which you want to spread into Data Part</param>
        /// <param name="aggregate">The Aggregate function which you want to apply in case matching data found more than once</param>
        /// <param name="columnFields">The List of column names which you want to spread as columns</param>
        /// <returns>A DataTable containing the Pivoted Data</returns>
        public DataTable PivotData(string rowField, string dataField, AggregateFunction aggregate, params string[] columnFields)
        {
            DataTable dt = new DataTable();
            string Separator = ".";
            List<string> rowList = _Source.Select(x => x[rowField].ToString()).Distinct().ToList();
            // Gets the list of columns .(dot) separated.
            var colList = _Source.Select(x => (columnFields.Select(n => x[n]).Aggregate((a, b) => a += Separator + b.ToString())).ToString()).Distinct().OrderBy(m => m);

            dt.Columns.Add(rowField);
            foreach (var colName in colList)
                dt.Columns.Add(colName);  // Cretes the result columns.//

            foreach (string rowName in rowList)
            {
                DataRow row = dt.NewRow();
                row[rowField] = rowName;
                foreach (string colName in colList)
                {
                    string strFilter = rowField + " = '" + rowName + "'";
                    string[] strColValues = colName.Split(Separator.ToCharArray(), StringSplitOptions.None);
                    for (int i = 0; i < columnFields.Length; i++)
                        strFilter += " and " + columnFields[i] + " = '" + strColValues[i] + "'";
                    if (colName != "")
                    {
                        row[colName] = GetData(strFilter, dataField, aggregate);
                    }
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

        public DataTable PivotData(string rowField, string dataField, AggregateFunction aggregate, bool showSubTotal, params string[] columnFields)
        {
            DataTable dt = new DataTable();
            string Separator = ".";
            List<string> rowList = _Source.Select(x => x[rowField].ToString()).Distinct().ToList();
            // Gets the list of columns .(dot) separated.
            List<string> colList = _Source.Select(x => columnFields.Aggregate((a, b) => x[a].ToString() + Separator + x[b].ToString())).Distinct().OrderBy(m => m).ToList();

            if (showSubTotal && columnFields.Length > 1)
            {
                string totalField = string.Empty;
                for (int i = 0; i < columnFields.Length - 1; i++)
                    totalField += columnFields[i] + "(Total)" + Separator;
                List<string> totalList = _Source.Select(x => totalField + x[columnFields.Last()].ToString()).Distinct().OrderBy(m => m).ToList();
                colList.InsertRange(0, totalList);
            }

            dt.Columns.Add(rowField);
            colList.ForEach(x => dt.Columns.Add(x));

            foreach (string rowName in rowList)
            {
                DataRow row = dt.NewRow();
                row[rowField] = rowName;
                foreach (string colName in colList)
                {
                    string filter = rowField + " = '" + rowName + "'";
                    string[] colValues = colName.Split(Separator.ToCharArray(), StringSplitOptions.None);
                    for (int i = 0; i < columnFields.Length; i++)
                        if (!colValues[i].Contains("(Total)"))
                            filter += " and " + columnFields[i] + " = '" + colValues[i] + "'";
                    row[colName] = GetData(filter, dataField, colName.Contains("(Total)") ? AggregateFunction.Sum : aggregate);
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

        public DataTable PivotData(string DataField, AggregateFunction Aggregate, string[] RowFields, string[] ColumnFields)
        {
            DataTable dt = new DataTable();
            string Separator = ".";
            var RowList = _SourceTable.DefaultView.ToTable(true, RowFields).AsEnumerable().ToList();
            for (int index = RowFields.Count() - 1; index >= 0; index--)
                RowList = RowList.OrderBy(x => x.Field<object>(RowFields[index])).ToList();
            // Gets the list of columns .(dot) separated.
            var ColList = (from x in _SourceTable.AsEnumerable()
                           select new
                           {
                               Name = ColumnFields.Select(n => x.Field<object>(n))
                                   .Aggregate((a, b) => a += Separator + b.ToString())
                           })
                               .Distinct()
                               .OrderBy(m => m.Name);

            //dt.Columns.Add(RowFields);
            foreach (string s in RowFields)
                dt.Columns.Add(s);

            foreach (var col in ColList)
                dt.Columns.Add(col.Name.ToString());  // Cretes the result columns.//

            foreach (var RowName in RowList)
            {
                DataRow row = dt.NewRow();
                string strFilter = string.Empty;

                foreach (string Field in RowFields)
                {
                    row[Field] = RowName[Field];
                    strFilter += " and " + Field + " = '" + RowName[Field].ToString() + "'";
                }
                strFilter = strFilter.Substring(5);

                foreach (var col in ColList)
                {
                    string filter = strFilter;
                    string[] strColValues = col.Name.ToString().Split(Separator.ToCharArray(), StringSplitOptions.None);
                    for (int i = 0; i < ColumnFields.Length; i++)
                        filter += " and " + ColumnFields[i] + " = '" + strColValues[i] + "'";
                    row[col.Name.ToString()] = GetData(filter, DataField, Aggregate);
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

        /// <summary>
        /// Retrives the data for matching RowField value and ColumnFields values with Aggregate function applied on them.
        /// </summary>
        /// <param name="Filter">DataTable Filter condition as a string</param>
        /// <param name="DataField">The column name which needs to spread out in Data Part of the Pivoted table</param>
        /// <param name="Aggregate">Enumeration to determine which function to apply to aggregate the data</param>
        /// <returns></returns>
        private object GetData(string Filter, string DataField, AggregateFunction Aggregate)
        {
            try
            {
                DataRow[] FilteredRows = _SourceTable.Select(Filter);
                object[] objList = FilteredRows.Select(x => x.Field<object>(DataField)).ToArray();

                switch (Aggregate)
                {
                    case AggregateFunction.Average:
                        return GetAverage(objList);
                    case AggregateFunction.Count:
                        return objList.Count();
                    case AggregateFunction.Exists:
                        return (objList.Count() == 0) ? "False" : "True";
                    case AggregateFunction.First:
                        return GetFirst(objList);
                    case AggregateFunction.Last:
                        return GetLast(objList);
                    case AggregateFunction.Max:
                        return GetMax(objList);
                    case AggregateFunction.Min:
                        return GetMin(objList);
                    case AggregateFunction.Sum:
                        return GetSum(objList);
                    default:
                        return null;
                }
            }
            catch (Exception ex)
            {;
                logger.Error(ex.Message + ":" + ex.StackTrace);
                return "#Error";
            }
        }

        private object GetAverage(object[] objList)
        {
            return objList.Count() == 0 ? null : (object)(Convert.ToDecimal(GetSum(objList)) / objList.Count());
        }
        private object GetSum(object[] objList)
        {
            return objList.Count() == 0 ? null : (object)(objList.Aggregate(new decimal(), (x, y) => x += Convert.ToDecimal(y)));
        }
        private object GetFirst(object[] objList)
        {
            return (objList.Count() == 0) ? null : objList.First();
        }
        private object GetLast(object[] objList)
        {
            return (objList.Count() == 0) ? null : objList.Last();
        }
        private object GetMax(object[] objList)
        {
            return (objList.Count() == 0) ? null : objList.Max();
        }
        private object GetMin(object[] objList)
        {
            return (objList.Count() == 0) ? null : objList.Min();
        }
    }

    public enum AggregateFunction
    {
        Count = 1,
        Sum = 2,
        First = 3,
        Last = 4,
        Average = 5,
        Max = 6,
        Min = 7,
        Exists = 8
    }
    #endregion
    public class PROJECT_TASK_TREE_NODE
    {
        public PLAN_TASK Task { get; set; }
        public LinkedList<PLAN_TASK> ChildTask { get; set; }
        public PLAN_TASK ParentTask { get; set; }
        public void addChild(PLAN_TASK childTask)
        {
            if (null == ChildTask)
            {
                ChildTask = new LinkedList<PLAN_TASK>();
            }
            ChildTask.AddLast(childTask);
        }
    }
    public class TASK_TREE4SHOW
    {
        public string text { get; set; }
        public string href { get; set; }
        public List<string> tags = new List<string>();
        //public TASK_TREE4SHOW ParentTask { get; set; }
        public LinkedList<TASK_TREE4SHOW> nodes { get; set; }
        public void addChild(TASK_TREE4SHOW childTask)
        {
            if (null == nodes)
            {
                nodes = new LinkedList<TASK_TREE4SHOW>();
            }
            nodes.AddLast(childTask);
        }
    }
    #region 公司組織
    public class DEPTARTMENT_TREE_NODE
    {
        public ENT_DEPARTMENT Dept { get; set; }
        public LinkedList<ENT_DEPARTMENT> ChildDept { get; set; }
        public ENT_DEPARTMENT ParentDept { get; set; }
        //加入下一級部門
        public void addChild(ENT_DEPARTMENT childdept)
        {
            if (null == ChildDept)
            {
                ChildDept = new LinkedList<ENT_DEPARTMENT>();
            }
            ChildDept.AddLast(childdept);
        }
    }
    public class DEPARTMENT_TREE4SHOW
    {
        public string text { get; set; }
        public string href { get; set; }
        public List<string> state;
        public List<string> tags = new List<string>();

        public LinkedList<DEPARTMENT_TREE4SHOW> nodes { get; set; }

        public void addChild(DEPARTMENT_TREE4SHOW childTask)
        {
            if (null == nodes)
            {
                nodes = new LinkedList<DEPARTMENT_TREE4SHOW>();
            }
            nodes.AddLast(childTask);
        }
    }
    #endregion

    #region Tree Stucture
    public abstract class TreeNode<T>
    {
        public T Value { get; set; }
        public abstract TreeNode<T> Parent { get; }
        public abstract TreeList<T> Children { get; }
        public abstract int Count { get; }
        public abstract int Degree { get; }
        public abstract int Depth { get; }
        public abstract int Level { get; }
        public TreeNode(T value)
        {
            this.Value = value;
        }
        public abstract void Add(T value);
        public abstract void Add(TreeNode<T> tree);
        public abstract void Remove();
        public abstract TreeNode<T> Clone();
    }

    public abstract class TreeList<T> : IEnumerable<TreeNode<T>>
    {
        public abstract int Count { get; }
        public abstract IEnumerator<TreeNode<T>> GetEnumerator();

        IEnumerator<TreeNode<T>> IEnumerable<TreeNode<T>>.GetEnumerator()
        {
            return GetEnumerator();
        }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }

    public class LinkedTree<T> : TreeNode<T>
    {
        protected LinkedList<LinkedTree<T>> childrenList;

        protected LinkedTree<T> parent;
        public override TreeNode<T> Parent
        {
            get
            {
                return parent;
            }
        }

        protected LinkedTreeList<T> children;
        public override TreeList<T> Children
        {
            get
            {
                return children;
            }
        }

        public override int Degree
        {
            get
            {
                return childrenList.Count;
            }
        }

        protected int count;
        public override int Count
        {
            get
            {
                return count;
            }
        }

        protected int depth;
        public override int Depth
        {
            get
            {
                return depth;
            }
        }

        protected int level;
        public override int Level
        {
            get
            {
                return level;
            }
        }

        public LinkedTree(T value) : base(value)
        {
            childrenList = new LinkedList<LinkedTree<T>>();
            children = new LinkedTreeList<T>(childrenList);
            depth = 1;
            level = 1;
            count = 1;
        }

        public override void Add(T value)
        {
            Add(new LinkedTree<T>(value));
        }

        public override void Add(TreeNode<T> tree)
        {
            LinkedTree<T> gtree = (LinkedTree<T>)tree;
            if (gtree.Parent != null)
                gtree.Remove();
            gtree.parent = this;
            if (gtree.depth + 1 > depth)
            {
                depth = gtree.depth + 1;
                BubbleDepth();
            }
            gtree.level = level + 1;
            gtree.UpdateLevel();
            childrenList.AddLast(gtree);
            count += tree.Count;
            BubbleCount(tree.Count);
        }

        public override void Remove()
        {
            if (parent == null)
                return;
            parent.childrenList.Remove(this);
            if (depth + 1 == parent.depth)
                parent.UpdateDepth();
            parent.count -= count;
            parent.BubbleCount(-count);
            parent = null;
        }

        public override TreeNode<T> Clone()
        {
            return Clone(1);
        }

        protected LinkedTree<T> Clone(int level)
        {
            LinkedTree<T> cloneTree = new LinkedTree<T>(Value);
            cloneTree.depth = depth;
            cloneTree.level = level;
            cloneTree.count = count;
            foreach (LinkedTree<T> child in childrenList)
            {
                LinkedTree<T> cloneChild = child.Clone(level + 1);
                cloneChild.parent = cloneTree;
                cloneTree.childrenList.AddLast(cloneChild);
            }
            return cloneTree;
        }

        protected void BubbleDepth()
        {
            if (parent == null)
                return;

            if (depth + 1 > parent.depth)
            {
                parent.depth = depth + 1;
                parent.BubbleDepth();
            }
        }

        protected void UpdateDepth()
        {
            int tmpDepth = depth;
            depth = 1;
            foreach (LinkedTree<T> child in childrenList)
                if (child.depth + 1 > depth)
                {
                    depth = child.depth + 1;
                }
            if (tmpDepth == depth || parent == null)
            {
                return;
            }
            if (tmpDepth + 1 == parent.depth)
            {
                parent.UpdateDepth();
            }
        }

        protected void BubbleCount(int diff)
        {
            if (parent == null)
                return;

            parent.count += diff;
            parent.BubbleCount(diff);
        }

        protected void UpdateLevel()
        {
            int childLevel = level + 1;
            foreach (LinkedTree<T> child in childrenList)
            {
                child.level = childLevel;
                child.UpdateLevel();
            }
        }
    }

    public class LinkedTreeList<T> : TreeList<T>
    {
        protected LinkedList<LinkedTree<T>> list;

        public LinkedTreeList(LinkedList<LinkedTree<T>> list)
        {
            this.list = list;
        }

        public override int Count
        {
            get
            {
                return list.Count;
            }
        }

        public override IEnumerator<TreeNode<T>> GetEnumerator()
        {
            return list.GetEnumerator();
        }
    }
    #endregion
}
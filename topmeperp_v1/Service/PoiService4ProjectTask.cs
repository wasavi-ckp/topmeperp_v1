using log4net;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    /// <summary>
    /// 解析專案任務與圖算匯入
    /// </summary>
    public class ProjectTask2MapService : ProjectItemFromExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public new string fileformat = "xlsx";
        public List<PLAN_TASK2MAPITEM> lstTask2Map = null;
        public ProjectTask2MapService()
        {
        }
        public void transAllSheet(string projectId)
        {
            //消防水
            lstTask2Map = new List<PLAN_TASK2MAPITEM>();
            try
            {
                logger.Debug("Process MapFW");
                lstTask2Map.AddRange(ConvertDataForMapFW(projectId));
                errorMessage = "消防水-匯入成功\\n";
            }
            catch(Exception ex)
            {
                logger.Error(ex.Message + ":" + ex.StackTrace);
                errorMessage = ex.Message;
            }
            //消防電
            try
            {
                logger.Debug("Process MapFP");
                lstTask2Map.AddRange(ConvertDataForMapFP(projectId));            
                errorMessage = errorMessage + ",消防電-匯入成功\\n";
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message + ":" + ex.StackTrace);
                errorMessage = errorMessage+ "," + ex.Message;
            }
            //電氣管線
            try
            {
                logger.Debug("Process MapPEP");
                lstTask2Map.AddRange(ConvertDataForMapPEP(projectId));
                errorMessage = errorMessage + ",電氣管線-匯入成功\\n";
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message + ":" + ex.StackTrace);
                errorMessage = errorMessage + ",電氣管線:" + ex.Message;
            }
            //給排水
            try
            {
                logger.Debug("Process MapPLU");
                lstTask2Map.AddRange(ConvertDataForMapPLU(projectId));
                errorMessage = errorMessage + ",給排水-匯入成功\\n";
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message + ":" + ex.StackTrace);
                errorMessage = errorMessage + ",給排水:" + ex.Message;
            }
        }
        #region 消防水資料轉換 
        /**
         * 取得消防水Sheet 資料
         * */
        public new List<PLAN_TASK2MAPITEM> ConvertDataForMapFW(string projectId)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + ":消防水");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("消防水");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + ":消防水");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("消防水");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有消防水資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[消防水]資料");
            }
            return ConverData2MapFW();
        }
        /**
         * 轉換圖算數量:消防水
         * */
        protected new List<PLAN_TASK2MAPITEM> ConverData2MapFW()
        {
            IRow row = null;
            List<PLAN_TASK2MAPITEM> lstTask2MapFw = new List<PLAN_TASK2MAPITEM>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..(表頭)
            while (iRowIndex < (1))
            {
                rows.MoveNext();
                iRowIndex++;
            }
            //循序處理每一筆資料之欄位!!
            iRowIndex++;
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                logger.Debug("Excel Value:" + row.Cells[0].ToString() + row.Cells[1] + row.Cells[2]);
                //將各Row 資料寫入物件內
                //0.項次	1.圖號	2.棟別	3.一次側位置	4.一次側名稱	5.二次側名稱	6.二次側位置	7.管材名稱	8管數/組	9管組數	10管長度/組數	11管總長
                if (row.Cells.Count <10 && row.Cells[0].ToString().ToUpper() != "END")
                {
                    lstTask2MapFw.AddRange(convertRow2TndMapFW(row, iRowIndex));
                }
                iRowIndex++;
            }
            logger.Info("MAP_FW Count:" + lstTask2MapFw.Count);
            return lstTask2MapFw;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件(消防水)
         * */
        private List<PLAN_TASK2MAPITEM> convertRow2TndMapFW(IRow row, int excelrow)
        {
            List<PLAN_TASK2MAPITEM> lstAllTask = new List<PLAN_TASK2MAPITEM>();
               
            if (row.Cells[7].ToString().Trim() != "" && row.Cells[8].ToString().Trim() != "")//0.專案認識識別碼
            {
                logger.Debug("create task map =" + row.Cells[7].ToString());
                string projectItemId = row.Cells[8].ToString().Trim();
                string[] strPrjUid = row.Cells[7].ToString().Split(',');
                for (int i = 0; i < strPrjUid.Length; i++)
                {
                    PLAN_TASK2MAPITEM item = new PLAN_TASK2MAPITEM();
                    item.PROJECT_ID = projId;
                    item.PRJ_UID = int.Parse(strPrjUid[i]);
                    logger.Debug("create task map =" + item.PRJ_UID);
                    item.PROJECT_ITEM_ID = projectItemId;
                    item.MAP_TYPE = "TND_MAP_FW";
                    item.MAP_PK = 0;
                    lstAllTask.Add(item);
                }
            }else
            {
                logger.Warn("Data format maybe have error(TND_MAP_FW):" + excelrow);
            }
            logger.Debug(excelrow + " had TND_MAP_FW Count:" + lstAllTask.Count);
            return lstAllTask;
        }
        #endregion
        #region 消防電資料轉換 
        /**
         ** 取得消防電Sheet 資料
         */
        public new List<PLAN_TASK2MAPITEM> ConvertDataForMapFP(string projectId)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + ":消防電");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("消防電");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + ":消防電");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("消防電");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有消防水資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[消防電]資料");
            }
            return ConverData2MapFP();
        }
        /**
         * 轉換圖算數量:消防電
         * */
        protected new List<PLAN_TASK2MAPITEM> ConverData2MapFP()
        {
            IRow row = null;
            List<PLAN_TASK2MAPITEM> lstTask2MapFp = new List<PLAN_TASK2MAPITEM>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..(表頭)
            while (iRowIndex < (1))
            {
                rows.MoveNext();
                iRowIndex++;
            }
            //循序處理每一筆資料之欄位!!
            iRowIndex++;
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                logger.Debug("Excel Value:" + row.Cells[0].ToString() + row.Cells[1] + row.Cells[2]);
                if (row.Cells.Count <10 && row.Cells[0].ToString().ToUpper() != "END")
                {
                    logger.Debug("FP iRows=" + iRowIndex);
                    lstTask2MapFp.AddRange(convertRow2TndMapFP(row, iRowIndex));
                }
                iRowIndex++;
            }
            logger.Info("MAP_FP Count:" + lstTask2MapFp.Count);
            return lstTask2MapFp;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件(消防電)
         * */
        private List<PLAN_TASK2MAPITEM> convertRow2TndMapFP(IRow row, int excelrow)
        {
            //消防電含有一筆管、一筆線資料
            List<PLAN_TASK2MAPITEM> lstItem = new List<PLAN_TASK2MAPITEM>();

            //消防電/線材
            if (row.Cells[7].ToString().Trim() != "" && row.Cells[8].ToString().Trim() != "")//0.專案認識識別碼
            {
                logger.Debug("crate task map =" + row.Cells[7].ToString());
                string[] strPrjUid = row.Cells[7].ToString().Split(',');
                for (int i = 0; i < strPrjUid.Length; i++)
                {
                    PLAN_TASK2MAPITEM item = new PLAN_TASK2MAPITEM();
                    item.PROJECT_ID = projId;
                    item.PRJ_UID = int.Parse(strPrjUid[i]);
                    item.PROJECT_ITEM_ID = row.Cells[8].ToString();
                    item.MAP_TYPE = "TND_MAP_FP";
                    item.MAP_PK = 0;
                    lstItem.Add(item);
                }
            }
            else
            {
                logger.Warn("Data format maybe have error(TND_MAP_FP):" + excelrow);
            }

            //消防電/管
            if (row.Cells[13].ToString().Trim() != "" && row.Cells[14].ToString().Trim() != "")//0.專案認識識別碼
            {
                logger.Debug("crate task map =" + row.Cells[13].ToString());
                string[] strPrjUid = row.Cells[13].ToString().Split(',');
                for (int i = 0; i < strPrjUid.Length; i++)
                {
                    PLAN_TASK2MAPITEM item = new PLAN_TASK2MAPITEM();
                    item.PROJECT_ID = projId;
                    item.PRJ_UID = int.Parse(strPrjUid[i]);
                    item.PROJECT_ITEM_ID = row.Cells[14].ToString();
                    item.MAP_TYPE = "TND_MAP_FP";
                    item.MAP_PK = 0;
                    lstItem.Add(item);
                }
            }
            else
            {
                logger.Warn("Data format maybe have error(TND_MAP_FP):" + excelrow);
            }
            logger.Debug(excelrow + " had TND_MAP_FP Count:" + lstItem.Count);
            return lstItem;
        }
        #endregion

        #region 電氣管線資料轉換 
        /**
         ** 取得電氣管線Sheet 資料
         */
        public new List<PLAN_TASK2MAPITEM> ConvertDataForMapPEP(string projectId)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + ":電氣管線");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("電氣管線");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + ":電氣管線");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("電氣管線");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有電氣管線資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[電氣管線]資料");
            }
            return ConverData2MapPEP();
        }
        /**
         * 轉換圖算數量:電氣管線
         * */
        protected new List<PLAN_TASK2MAPITEM> ConverData2MapPEP()
        {
            IRow row = null;
            List<PLAN_TASK2MAPITEM> lstTask2MapPep = new List<PLAN_TASK2MAPITEM>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..(表頭)
            while (iRowIndex < (1))
            {
                rows.MoveNext();
                iRowIndex++;
            }
            //循序處理每一筆資料之欄位!!
            iRowIndex++;
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                logger.Debug("PEP Excel Value:" + iRowIndex+ ","+ row.Cells.Count);
                if (row.Cells.Count > 10 && row.Cells[0].ToString().ToUpper() != "END")
                {
                    logger.Debug("PEP iRows=" + iRowIndex);
                    lstTask2MapPep.AddRange(convertRow2TndMapPEP(row, iRowIndex));
                }
                iRowIndex++;
            }
            logger.Info("MAP_PEP Count:" + lstTask2MapPep.Count);
            return lstTask2MapPep;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件-電氣管線()
         * */
        private List<PLAN_TASK2MAPITEM> convertRow2TndMapPEP(IRow row, int excelrow)
        {
            //電氣管線含有一筆管、一筆線資料、一筆地線資料
            List<PLAN_TASK2MAPITEM> lstItem = new List<PLAN_TASK2MAPITEM>();

            //電/線材
            if (row.Cells[7].ToString().Trim() != "" && row.Cells[8].ToString().Trim() != "")//0.專案認識識別碼
            {
                logger.Debug("crate task map =" + row.Cells[7].ToString());
                string[] strPrjUid = row.Cells[7].ToString().Split(',');
                for (int i = 0; i < strPrjUid.Length; i++)
                {
                    PLAN_TASK2MAPITEM item = new PLAN_TASK2MAPITEM();
                    item.PROJECT_ID = projId;
                    item.PRJ_UID = int.Parse(strPrjUid[i]);
                    item.PROJECT_ITEM_ID = row.Cells[8].ToString();
                    item.MAP_TYPE = "TND_MAP_PEP";
                    item.MAP_PK = 0;
                    lstItem.Add(item);
                }
            }
            else
            {
                logger.Warn("Data format maybe have error(TND_MAP_PEP):" + excelrow);
            }

            //電地線
            if (row.Cells[13].ToString().Trim() != "" && row.Cells[14].ToString().Trim() != "")//0.專案認識識別碼
            {
                logger.Debug("crate task map =" + row.Cells[13].ToString());
                string[] strPrjUid = row.Cells[13].ToString().Split(',');
                for (int i = 0; i < strPrjUid.Length; i++)
                {
                    PLAN_TASK2MAPITEM item = new PLAN_TASK2MAPITEM();
                    item.PROJECT_ID = projId;
                    item.PRJ_UID = int.Parse(strPrjUid[i]);
                    item.PROJECT_ITEM_ID = row.Cells[14].ToString();
                    item.MAP_TYPE = "TND_MAP_PEP";
                    item.MAP_PK = 0;
                    lstItem.Add(item);
                }
            }
            else
            {
                logger.Warn("Data format maybe have error(TND_MAP_PEP):" + excelrow);
            }

            //電氣管
            if (row.Cells[17].ToString().Trim() != "" && row.Cells[18].ToString().Trim() != "")//0.專案認識識別碼
            {
                logger.Debug("crate task map =" + row.Cells[17].ToString());
                string[] strPrjUid = row.Cells[17].ToString().Split(',');
                for (int i = 0; i < strPrjUid.Length; i++)
                {
                    PLAN_TASK2MAPITEM item = new PLAN_TASK2MAPITEM();
                    item.PROJECT_ID = projId;
                    item.PRJ_UID = int.Parse(strPrjUid[i]);
                    item.PROJECT_ITEM_ID = row.Cells[18].ToString();
                    item.MAP_TYPE = "TND_MAP_PEP";
                    item.MAP_PK = 0;
                    lstItem.Add(item);
                }
            }
            else
            {
                logger.Warn("Data format maybe have error(TND_MAP_PEP):" + excelrow);
            }
            logger.Debug(excelrow + " had TND_MAP_PEP Count:" + lstItem.Count);

            return lstItem;
        }
        #endregion
        #region 給排水資料轉換 
        /**
         ** 取得給排水Sheet 資料
         */
        public new List<PLAN_TASK2MAPITEM> ConvertDataForMapPLU(string projectId)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + ":給排水");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("給排水");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + ":給排水");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("給排水");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有給排水資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[給排水]資料");
            }
            return ConverData2MapPLU();
        }
        /**
         * 轉換圖算數量:給排水
         * */
        protected new List<PLAN_TASK2MAPITEM> ConverData2MapPLU()
        {
            IRow row = null;
            List<PLAN_TASK2MAPITEM> lstTask2MapPlu = new List<PLAN_TASK2MAPITEM>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..(表頭)
            while (iRowIndex < (1))
            {
                rows.MoveNext();
                iRowIndex++;
            }
            //循序處理每一筆資料之欄位!!
            iRowIndex++;
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                logger.Debug("PLU Excel Value:" + iRowIndex + "," + row.Cells.Count);
                if (row.Cells.Count >7 && row.Cells[0].ToString().ToUpper() != "END")
                {
                    logger.Debug("PEP iRows=" + iRowIndex);
                    lstTask2MapPlu.AddRange(convertRow2TndMapPLU(row, iRowIndex));
                }
                iRowIndex++;
            }
            logger.Info("MAP_PEP Count:" + lstTask2MapPlu.Count);
            return lstTask2MapPlu;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件-給排水()
         * */
        private List<PLAN_TASK2MAPITEM> convertRow2TndMapPLU(IRow row, int excelrow)
        {
            //給排水含有一筆管
            List<PLAN_TASK2MAPITEM> lstItem = new List<PLAN_TASK2MAPITEM>();

            //給排水
            if (row.Cells[7].ToString().Trim() != "" && row.Cells[8].ToString().Trim() != "")//0.專案認識識別碼
            {
                logger.Debug("crate task map =" + row.Cells[7].ToString());
                string[] strPrjUid = row.Cells[7].ToString().Split(',');
                for (int i = 0; i < strPrjUid.Length; i++)
                {
                    PLAN_TASK2MAPITEM item = new PLAN_TASK2MAPITEM();
                    item.PROJECT_ID = projId;
                    item.PRJ_UID = int.Parse(strPrjUid[i]);
                    item.PROJECT_ITEM_ID = row.Cells[8].ToString();
                    item.MAP_TYPE = "TND_MAP_PLU";
                    item.MAP_PK = 0;
                    lstItem.Add(item);
                }
            }
            else
            {
                logger.Warn("Data format maybe have error(TND_MAP_PLU):" + excelrow);
            }

            logger.Debug(excelrow + " had TND_MAP_PLU Count:" + lstItem.Count);

            return lstItem;
        }
        #endregion
    }
}
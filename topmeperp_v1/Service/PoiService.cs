using log4net;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    //提供Excle 樣板檔案基本做法
    public class ExcelBase
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public IWorkbook hssfworkbook = null;
        public static string strUploadPath = ConfigurationManager.AppSettings["UploadFolder"];
        public ISheet sheet = null;
        public string fileformat = "xlsx";
        //定義樣板檔案名稱
        public string templateFile = null;
        //定義輸出檔案名稱
        public string outputFile = null;
        public string errMessage = null;
        public XSSFCellStyle style = null;
        public XSSFCellStyle styleNumber = null;

        public ExcelBase()
        {

        }

        public void DataTableToExcelFile(DataTable dt, int startRowIndex = 0)
        {
            //建立Excel 2003檔案
            hssfworkbook = new HSSFWorkbook();

            ////建立Excel 2007檔案
            //hssfworkbook = new XSSFWorkbook();
            SetDataTabletoSheet(dt, startRowIndex);

            FileStream file = new FileStream(outputFile, FileMode.Create);//產生檔案
            hssfworkbook.Write(file);
            file.Close();
        }
        protected void SetDataTabletoSheet(DataTable dt, int startRowIndex)
        {
            if (dt.TableName != string.Empty)
            {
                sheet = hssfworkbook.CreateSheet(dt.TableName);
            }
            else
            {
                sheet = hssfworkbook.CreateSheet("Sheet1");
            }

            sheet.CreateRow(startRowIndex);//第一行為欄位名稱
            for (int i = startRowIndex; i < dt.Columns.Count; i++)
            {
                sheet.GetRow(startRowIndex).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }

            for (int i = startRowIndex; i < dt.Rows.Count; i++)
            {
                sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    sheet.GetRow(i + 1).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
                }
            }
        }
        public byte[] DataTableToExcelFileByResponse(DataTable dt, int startRowIndex = 0)
        {
            ////建立Excel 2007檔案
            hssfworkbook = new XSSFWorkbook();
            SetDataTabletoSheet(dt, startRowIndex);
            MemoryStream ms = new MemoryStream();
            hssfworkbook.Write(ms);
            byte[] msByte = ms.ToArray();
            ms.Close();
            ms.Dispose();
            return msByte;
        }
        /*讀取Excel 樣板檔案!!!*/
        public void InitializeWorkbook()
        {
            //read the template via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
            //book1.xls is an Excel-2007-generated file, so some new unknown BIFF records are added. 
            using (FileStream file = new FileStream(templateFile, FileMode.Open, FileAccess.Read))
            {
                logger.Info("Read Excel File:" + templateFile);
                if (file.Name.EndsWith(".xls"))
                {
                    logger.Debug("process excel file for office 2003");
                    fileformat = "xls";
                    hssfworkbook = new HSSFWorkbook(file);
                }
                else
                {
                    logger.Debug("process excel file for office 2007");
                    hssfworkbook = new XSSFWorkbook(file);
                }
                file.Close();
            }
            style = ExcelStyle.getContentStyle(hssfworkbook);
            styleNumber = ExcelStyle.getNumberStyle(hssfworkbook);
        }
        //設定讀取之Excel 檔案
        public void InitializeWorkbook(string path)
        {
            //read the template via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
            //book1.xls is an Excel-2007-generated file, so some new unknown BIFF records are added. 
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                logger.Info("Read Excel File:" + path); if (file.Name.EndsWith(".xls"))
                {
                    logger.Debug("process excel file for office 2003");
                    fileformat = "xls";
                    hssfworkbook = new HSSFWorkbook(file);
                }
                else
                {
                    logger.Debug("process excel file for office 2007");
                    hssfworkbook = new XSSFWorkbook(file);
                }
                file.Close();
            }
        }
        //Initial Sheet
        public void SetOpSheet(string sheetName)
        {
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat);
                sheet = (HSSFSheet)hssfworkbook.GetSheet(sheetName);
            }
            else
            {
                logger.Debug("office 2007:" + fileformat);
                sheet = (XSSFSheet)hssfworkbook.GetSheet(sheetName);
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有Sheet=" + sheetName + "! filename=" + fileformat);
                throw new Exception("檔案內沒有Sheet:" + sheetName);
            }
        }
    }

    /// <summary>
    /// 備標階段標單品項清單
    /// </summary>
    public class ProjectItemFromExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public IWorkbook hssfworkbook;
        public ISheet sheet = null;
        protected string fileformat = "xlsx";
        protected string projId = null;
        public List<TND_PROJECT_ITEM> lstProjectItem = null;
        public List<PLAN_ITEM> lstPlanItem = null;
        public string errorMessage = null;
        //test conflicts
        public ProjectItemFromExcel()
        {
        }
        /*讀取備標Excel 檔案!!!*/
        public void InitializeWorkbook(string path)
        {
            //read the template via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
            //book1.xls is an Excel-2007-generated file, so some new unknown BIFF records are added. 
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                logger.Info("Read Excel File:" + path); if (file.Name.EndsWith(".xls"))
                {
                    logger.Debug("process excel file for office 2003");
                    fileformat = "xls";
                    hssfworkbook = new HSSFWorkbook(file);
                }
                else
                {
                    logger.Debug("process excel file for office 2007");
                    hssfworkbook = new XSSFWorkbook(file);
                }
                file.Close();
            }
        }
        //處理整理後標單..
        public void ConvertDataForTenderProject(string projectId, int startrow)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId);
                sheet = (HSSFSheet)hssfworkbook.GetSheet("整理後標單");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId);
                sheet = (XSSFSheet)hssfworkbook.GetSheet("整理後標單");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有整理後標單資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有整理後標單資料");
            }
            ConvertExcelToTndProjectItem(startrow);
        }
        //轉換標單內容物件
        public void ConvertExcelToTndProjectItem(int startrow)
        {
            IRow row = null;
            lstProjectItem = new List<TND_PROJECT_ITEM>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..
            while (iRowIndex < (startrow - 1))
            {
                iRowIndex++;
                //錯誤控制，補足 Row=null & Cells[0] ==null 情況以利資料匯入
                row = (IRow)sheet.GetRow(iRowIndex);
                if (null == row)
                {
                    logger.Debug("Empty Row Index=" + iRowIndex);
                }
                else if (row.Cells.Count > 2)
                {
                    logger.Debug("skip data Excel Value:" + iRowIndex + ","
                        + row.GetCell(0) == null ? "not cell" : row.Cells[0].ToString() + ","
                        + row.GetCell(1) == null ? "not cell" : row.Cells[1].ToString() + ","
                        + row.GetCell(2) == null ? "not cell" : row.Cells[2].ToString());
                }
                else
                {
                    logger.Debug("Index=" + iRowIndex);
                }
            }
            //循序處理每一筆資料之欄位!!
            //iRowIndex++;
            int itemId = 1;
            bool hasMore = true;
            while (hasMore)
            {
                row = (IRow)sheet.GetRow(iRowIndex);
                try
                {
                    if (row != null)
                    {
                        if (null == row.GetCell(0))
                        {
                            row.CreateCell(0).SetCellValue("");
                        }
                        logger.Debug("Cells Count=" + row.Cells.Count + ",Excel Index:" + iRowIndex);
                    }
                    else
                    {
                        logger.Warn("empty row:" + iRowIndex);
                        row = sheet.CreateRow(iRowIndex);
                        row.CreateCell(0).SetCellValue("");
                    }
                    //將各Row 資料寫入物件內
                    //項次,名稱,單位,數量,單價,複價,備註,九宮格,次九宮格,主系統,次系統
                    if (row.GetCell(0).ToString().ToUpper().Trim() == "END")
                    {
                        logger.Info("Finish convert Job : count=" + lstProjectItem.Count);
                        hasMore = false;
                        return;
                    }
                    else
                    {
                        lstProjectItem.Add(convertRow2TndProjectItem(itemId, row, iRowIndex));
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(iRowIndex + ",ex:" + ex.Message + "," + ex.StackTrace);
                }
                iRowIndex++;
                itemId++;
            }
        }
        private TND_PROJECT_ITEM convertRow2TndProjectItem(int id, IRow row, int excelrow)
        {
            TND_PROJECT_ITEM projectItem = new TND_PROJECT_ITEM();
            projectItem.PROJECT_ID = projId;
            logger.Debug("irow=" + excelrow + ",Row=" + row + "," + row.Count());

            if (row.GetCell(0) != null && row.GetCell(0).ToString().Trim() != "")//項次
            {
                projectItem.ITEM_ID = row.GetCell(0).ToString();
            }
            if (row.GetCell(1) != null && row.GetCell(1).ToString().Trim() != "")//名稱
            {
                projectItem.ITEM_DESC = row.GetCell(1).ToString();
            }
            if (row.GetCell(2) != null && row.GetCell(2).ToString().Trim() != "")//單位
            {
                projectItem.ITEM_UNIT = row.GetCell(2).ToString();
            }
            if (row.GetCell(3) != null && row.GetCell(3).ToString().Trim() != "")//數量
            {
                try
                {
                    decimal dQty = decimal.Parse(row.GetCell(3).ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.GetCell(3).ToString());
                    projectItem.ITEM_QUANTITY = dQty;
                }
                catch (Exception e)
                {
                    logger.Error("data format Error on ExcelRow=" + excelrow + ",Item_Desc= " + projectItem.ITEM_DESC + ",value=" + row.GetCell(3).ToString());
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Item_Desc= " + projectItem.ITEM_DESC + ",value=" + row.GetCell(3).ToString());
                    logger.Error(e.Message);
                }

            }
            if (row.GetCell(6) != null && row.GetCell(6).ToString().Trim() != "")//備註
            {
                projectItem.ITEM_REMARK = row.GetCell(6).ToString();
            }
            if (row.GetCell(7) != null && row.GetCell(7).ToString().Trim() != "")//九宮格
            {
                projectItem.TYPE_CODE_1 = row.GetCell(7).ToString();
            }
            if (row.GetCell(8) != null && row.GetCell(8).ToString().Trim() != "")//次九宮格
            {
                projectItem.TYPE_CODE_2 = row.GetCell(8).ToString();
            }
            if (row.GetCell(9) != null && row.GetCell(9).ToString().Trim() != "")//主系統
            {
                projectItem.SYSTEM_MAIN = row.GetCell(9).ToString();
            }
            if (row.GetCell(10) != null && row.GetCell(10).ToString().Trim() != "")//次系統
            {
                projectItem.SYSTEM_SUB = row.GetCell(10).ToString();
            }
            projectItem.PROJECT_ITEM_ID = projId + "-" + id;
            projectItem.EXCEL_ROW_ID = excelrow;
            projectItem.CREATE_DATE = System.DateTime.Now;

            logger.Info("TndprojectItem=" + projectItem.ToString());
            return projectItem;
        }

        #region 處理得標標單項目內容
        //處理得標後標單..
        public void ConvertDataForPlan(string projectId)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟得標後標單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId);
                sheet = (HSSFSheet)hssfworkbook.GetSheet("合約標單");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId);
                sheet = (XSSFSheet)hssfworkbook.GetSheet("合約標單");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有得標後標單資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有合約標單資料");
            }
            ConvertExcelToPlanItem();
        }
        //轉換標單內容物件
        public void ConvertExcelToPlanItem()
        {
            IRow row = null;
            lstPlanItem = new List<PLAN_ITEM>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..
            while (iRowIndex < 4)
            {
                rows.MoveNext();
                iRowIndex++;
                row = (IRow)rows.Current;
                logger.Debug("skip data Excel Value:" + row.Cells[0].ToString() + "," + row.Cells[1] + "," + row.Cells[2]);
            }
            //循序處理每一筆資料之欄位!!
            iRowIndex++;
            int itemId = 1;
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                logger.Debug("Cells Count=" + row.Cells.Count + ",Excel Value:" + row.Cells[0].ToString());
                //將各Row 資料寫入物件內
                //代碼,項次,名稱,單位,標單數量,採購數量,單價,複價,備註,九宮格,次九宮格,主系統,次系統
                if (row.Cells[0].ToString().ToUpper() != "" && row.Cells[0].ToString().ToUpper() != "END")//代碼欄位有值才匯入
                {
                    lstPlanItem.Add(convertRow2PlanItem(itemId, row, iRowIndex));
                }
                else
                {
                    logger.Info("Finish convert Job : count=" + lstPlanItem.Count);
                    return;
                }
                iRowIndex++;
                itemId++;
            }
        }
        private PLAN_ITEM convertRow2PlanItem(int id, IRow row, int excelrow)
        {
            logger.Debug("Cells Count=" + row.Cells.Count + ",Row PhysicalNumberOfCells no=:" + row.PhysicalNumberOfCells);
            PLAN_ITEM planItem = new PLAN_ITEM();
            planItem.PROJECT_ID = projId;
            planItem.EXCEL_ROW_ID = excelrow;
            planItem.CREATE_DATE = System.DateTime.Now;

            ///改用迴圈處理每個Cell
            for (int i = 0; i < 15; i++)
            {
                /// Sample
                ///  If row.GetCell(j).CellType = CellType.FORMULA Then    '== v.1.2.4版修改
                //      D_dataRow(j) = row.GetCell(j).NumericCellValue
                //    '-- 表示格子裡面，公式運算後的「值」，是數字（Numeric）。而非抓到「公式」。
                //    Else
                //        D_dataRow(j) = row.GetCell(j).StringCellValue   '--每一個欄位，都加入同一列 DataRow
                //    End If
                ICell cell = row.GetCell(i);
                try
                {
                    logger.Debug("for feaure rowID=" + id + ",cell style=" + row.GetCell(i).CellType);
                }
                catch (Exception ex)
                {
                    logger.Warn("for feaure rowID=" + id + ",cell tyle=" + i + "ex:" + ex.Message);
                }

                if (null != cell)
                {
                    switch (i)
                    {
                        case 0:
                            //代碼
                            planItem.PLAN_ITEM_ID = cell.ToString();
                            break;
                        case 1:
                            planItem.ITEM_ID = cell.ToString();
                            break;
                        case 2://名稱
                            planItem.ITEM_DESC = cell.ToString();
                            break;
                        case 3:
                            planItem.ITEM_UNIT = cell.ToString();
                            break;
                        case 4:
                            try
                            {
                                decimal dQty = decimal.Parse(cell.ToString());
                                logger.Info("excelrow=" + excelrow + ",value=" + cell.ToString());
                                planItem.ITEM_QUANTITY = dQty;
                            }
                            catch (Exception e)
                            {
                                logger.Error("data format Error on PLAN_ITEM_ID= " + planItem.PLAN_ITEM_ID + ",ExcelRow=" + excelrow + ",Item_Desc= " + planItem.ITEM_DESC + ",value=" + cell.ToString());
                                logErrorMessage("data format Error on PLAN_ITEM_ID= " + planItem.PLAN_ITEM_ID + ",ExcelRow=" + excelrow + ",Item_Desc= " + planItem.ITEM_DESC + ",value=" + cell.ToString());
                                logger.Error(e.Message + "," + e.StackTrace);
                            }

                            break;
                        case 5:
                            try
                            {
                                decimal dQty = decimal.Parse(cell.ToString());
                                logger.Info("excelrow=" + excelrow + ",value=" + cell.ToString());
                                planItem.ITEM_FORM_QUANTITY = dQty;
                            }
                            catch (Exception e)
                            {
                                logger.Error("data format Error on PLAN_ITEM_ID= " + planItem.PLAN_ITEM_ID + ",ExcelRow=" + excelrow + ",Item_Desc= " + planItem.ITEM_DESC + ",value=" + cell.ToString());
                                logErrorMessage("data format Error on PLAN_ITEM_ID= " + planItem.PLAN_ITEM_ID + ",ExcelRow=" + excelrow + ",Item_Desc= " + planItem.ITEM_DESC + ",value=" + cell.ToString());
                                logger.Error(e.Message);
                            }

                            break;
                        case 6://項目標單單價
                            try
                            {
                                decimal dQty = decimal.Parse(cell.ToString());
                                logger.Info("excelrow=" + excelrow + ",value=" + dQty);
                                planItem.ITEM_UNIT_PRICE = dQty;
                            }
                            catch (Exception e)
                            {
                                logger.Error("data format Error on PLAN_ITEM_ID= " + planItem.PLAN_ITEM_ID + ",ExcelRow=" + excelrow + ",Item_Desc= " + planItem.ITEM_DESC + ",value=" + cell.ToString());
                                logErrorMessage("data format Error on PLAN_ITEM_ID= " + planItem.PLAN_ITEM_ID + ",ExcelRow=" + excelrow + ",Item_Desc= " + planItem.ITEM_DESC + ",value=" + cell.ToString());
                                logger.Error(e.Message);
                            }
                            break;
                        case 7:  //複價欄位Skip
                            break;
                        case 8://備註
                            planItem.ITEM_REMARK = cell.ToString();
                            break;
                        case 9:
                            logger.Debug(id + " i index=" + i + "=" + cell.ToString());
                            planItem.TYPE_CODE_1 = cell.ToString();
                            break;
                        case 10:
                            planItem.TYPE_CODE_2 = cell.ToString();
                            break;
                        case 11:
                            planItem.SYSTEM_MAIN = cell.ToString();
                            break;
                        case 12:
                            planItem.SYSTEM_SUB = cell.ToString();
                            break;
                        case 13:
                            planItem.EXCEL_ROW_ID = Convert.ToInt32(cell.ToString());
                            break;
                        case 14:
                            planItem.IN_CONTRACT = cell.ToString();
                            break;
                    }
                }
            }
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(planItem);
            logger.Info("PlanItem=" + itemJson);
            return planItem;
        }
        #endregion
        #region 消防水資料轉換 
        /**
         * 取得消防水Sheet 資料
         * */
        public List<TND_MAP_FW> ConvertDataForMapFW(string projectId)
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
        protected List<TND_MAP_FW> ConverData2MapFW()
        {
            IRow row = null;
            List<TND_MAP_FW> lstMapFW = new List<TND_MAP_FW>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..(表頭)
            while (iRowIndex < (1))
            {
                rows.MoveNext();
                iRowIndex++;
                //row = (IRow)rows.Current;
                //logger.Debug("skip data Excel Value:" + row.Cells[0].ToString() + "," + row.Cells[1] + "," + row.Cells[2]);
            }
            //循序處理每一筆資料之欄位!!
            iRowIndex++;
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                //將各Row 資料寫入物件內
                //0.項次	1.圖號	2.棟別	3.一次側位置	4.一次側名稱	5.二次側名稱	6.二次側位置	7.管材名稱	8管數/組	9管組數	10管長度/組數	11管總長
                if (row.Cells.Count != 0 && row.Cells[0].ToString().ToUpper() != "END")
                {
                    logger.Debug("ROW:" + iRowIndex + ",Excel Value:" + row.Cells[0].ToString() + row.Cells[1] + row.Cells[2]);
                    lstMapFW.Add(convertRow2TndMapFW(row, iRowIndex));
                }
                else
                {
                    logErrorMessage("Step1 ;取得圖算數量消防水:" + lstMapFW.Count + "筆");
                    logger.Info("Finish convert Job : count=" + lstMapFW.Count);
                    return lstMapFW;
                }
                iRowIndex++;
            }
            logger.Info("MAP_FW Count:" + iRowIndex);
            return lstMapFW;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件
         * */
        private TND_MAP_FW convertRow2TndMapFW(IRow row, int excelrow)
        {
            logger.Debug("get last row no=" + row.LastCellNum);
            TND_MAP_FW item = new TND_MAP_FW();
            item.PROJECT_ID = projId;
            if (row.Cells[0].ToString().Trim() != "")//0.項次
            {
                item.EXCEL_ITEM = row.Cells[0].ToString();
            }
            if (row.Cells[1].ToString().Trim() != "")//1.圖號
            {
                item.MAP_NO = row.Cells[1].ToString();
            }
            if (row.Cells[2].ToString().Trim() != "")//2.棟別
            {
                item.BUILDING_NO = row.Cells[2].ToString();
            }

            if (row.Cells[3].ToString().Trim() != "")//3.一次側位置
            {
                item.PRIMARY_SIDE = row.Cells[3].ToString();
            }
            if (row.Cells[4].ToString().Trim() != "")//4.一次側名稱
            {
                item.PRIMARY_SIDE_NAME = row.Cells[4].ToString();
            }

            if (row.Cells[5].ToString().Trim() != "")//5.二次側名稱
            {
                item.SECONDARY_SIDE = row.Cells[5].ToString();
            }
            if (row.Cells[6].ToString().Trim() != "")//6.二次側位置
            {
                item.SECONDARY_SIDE_NAME = row.Cells[6].ToString();
            }

            if (row.Cells[7].ToString().Trim() != "")//7.管材名稱
            {
                item.PIPE_NAME = row.Cells[7].ToString();
            }

            if (null != row.Cells[8].ToString().Trim() || row.Cells[8].ToString().Trim() != "")// 8管數/組
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[8].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[8].ToString());
                    item.PIPE_CNT = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[8].value=" + row.Cells[8].ToString());
                    logger.Error(e.Message);
                }
            }

            if (null != row.Cells[9].ToString().Trim() || row.Cells[9].ToString().Trim() != "")// 9.管組數	
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[9].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[9].ToString());
                    item.PIPE_SET = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[9].value=" + row.Cells[9].ToString());
                    //   logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Item_Desc= " + projectItem.ITEM_DESC + ",value=" + row.Cells[3].ToString());
                    logger.Error(e.Message);
                }
            }

            if (null != row.Cells[10].ToString().Trim() || row.Cells[10].ToString().Trim() != "")// 10管長度/組數
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[10].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[10].ToString());
                    item.PIPE_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[10].value=" + row.Cells[10].ToString());
                    logger.Error(e.Message);
                }
            }

            string col11 = row.Cells[11].ToString().Trim();
            if (null != row.Cells[11].ToString().Trim() || row.Cells[11].ToString().Trim() != "")// 11管總長
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[11].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[11].ToString());
                    item.PIPE_TOTAL_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[11].value=" + row.Cells[11].ToString());
                    logger.Error(e.Message);
                }
            }
            item.CREATE_DATE = System.DateTime.Now;
            logger.Info("TND_MAP_FW=" + item.ToString());
            return item;
        }
        #endregion
        #region 消防電資料轉換 
        // * 處理消防電，進入點
        public List<TND_MAP_FP> ConvertDataForMapFP(string projectId)
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
                logger.Error("檔案內沒有消防電資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[消防電]資料");
            }
            return ConverData2MapFP();
        }
        /**
         * 轉換圖算數量:消防電
         * */
        protected List<TND_MAP_FP> ConverData2MapFP()
        {
            IRow row = null;
            List<TND_MAP_FP> lstMapFP = new List<TND_MAP_FP>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..(表頭)
            while (iRowIndex < 1)
            {
                rows.MoveNext();
                iRowIndex++;
                //row = (IRow)rows.Current;
                //logger.Debug("skip data Excel Value:" + row.Cells[0].ToString() + "," + row.Cells[1] + "," + row.Cells[2]);
            }
            //循序處理每一筆資料之欄位!!
            iRowIndex++;
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                //將各Row 資料寫入物件內
                //1.項次	2.圖號	3.棟別	4.一次側位置	5.一次側名稱	6.二次側名稱	
                //7.二次側位置	8.線材名稱	9.條數/組	10.線組數	11.線長度/條數	12.線總長	
                //13.管材名稱	14.管長	15.管組數	16.管總長
                if (row.Cells.Count != 0 && row.Cells[0].ToString().ToUpper() != "END")
                {
                    logger.Debug("Excel Value:" + row.Cells[0].ToString() + row.Cells[1] + row.Cells[2]);
                    lstMapFP.Add(convertRow2TndMapFP(row, iRowIndex));
                }
                else
                {
                    logErrorMessage("Step1 ;取得圖算數量消防電:" + lstMapFP.Count + "筆");
                    logger.Info("Finish convert Job : count=" + lstMapFP.Count);
                    return lstMapFP;
                }
                iRowIndex++;
            }
            logger.Info("MAP_FP Count:" + iRowIndex);
            return lstMapFP;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件
         * */
        private TND_MAP_FP convertRow2TndMapFP(IRow row, int excelrow)
        {
            TND_MAP_FP item = new TND_MAP_FP();
            item.PROJECT_ID = projId;
            //0.項次	1.圖號	2.棟別	3.一次側位置 4.一次側名稱	5.二次側名稱	

            if (row.Cells[0].ToString().Trim() != "")//0.項次
            {
                item.EXCEL_ITEM = row.Cells[0].ToString();
            }
            if (row.Cells[1].ToString().Trim() != "")//1.圖號
            {
                item.MAP_NO = row.Cells[1].ToString();
            }
            if (row.Cells[2].ToString().Trim() != "")//2.棟別
            {
                item.BUILDING_NO = row.Cells[2].ToString();
            }

            if (row.Cells[3].ToString().Trim() != "")//3.一次側位置
            {
                item.PRIMARY_SIDE = row.Cells[3].ToString();
            }
            if (row.Cells[4].ToString().Trim() != "")//4.一次側名稱
            {
                item.PRIMARY_SIDE_NAME = row.Cells[4].ToString();
            }

            if (row.Cells[5].ToString().Trim() != "")//5.二次側名稱
            {
                item.SECONDARY_SIDE = row.Cells[5].ToString();
            }
            //6.二次側位置	7.線材名稱	8.條數/組
            if (row.Cells[6].ToString().Trim() != "")//6.二次側位置
            {
                item.SECONDARY_SIDE_NAME = row.Cells[6].ToString();
            }

            if (row.Cells[7].ToString().Trim() != "")//7.線材名稱
            {
                item.WIRE_NAME = row.Cells[7].ToString();
            }

            if (row.Cells[8].ToString().Trim() != "")// 8.條數/組
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[8].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[8].ToString());
                    item.WIRE_QTY_SET = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[8].value=" + row.Cells[8].ToString());
                    logger.Error(e.Message);
                }
            }
            //9.線組數   10.線長度 / 條數   11.線總長      
            if (row.Cells[9].ToString().Trim() != "")// 9.線組數
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[9].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[9].ToString());
                    item.WIRE_SET_CNT = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[9].value=" + row.Cells[9].ToString());
                    //   logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Item_Desc= " + projectItem.ITEM_DESC + ",value=" + row.Cells[3].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[10].ToString().Trim() != "")// 10線長度 / 條數
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[10].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[10].ToString());
                    item.WIRE_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[11].value=" + row.Cells[10].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[11].ToString().Trim() != "")// 11.線總長
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[11].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[11].ToString());
                    item.WIRE_TOTAL_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[11].value=" + row.Cells[11].ToString());
                    logger.Error(e.Message);
                }
            }
            //12.管材名稱	13.管長	14.管組數	15.管總長
            if (row.Cells[12].ToString().Trim() != "")// 12.管材名稱
            {
                item.PIPE_NAME = row.Cells[12].ToString();
            }

            if (row.Cells[13].ToString().Trim() != "")//13.管長
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[10].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[13].ToString());
                    item.PIPE_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[13].value=" + row.Cells[10].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[14].ToString().Trim() != "")//14.管組數	
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[11].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[14].ToString());
                    item.PIPE_SET = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[14].value=" + row.Cells[11].ToString());
                    logger.Error(e.Message);
                }
            }
            if (row.Cells[15].ToString().Trim() != "")//15.管組數	
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[15].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[15].ToString());
                    item.PIPE_TOTAL_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[15].value=" + row.Cells[11].ToString());
                    logger.Error(e.Message);
                }
            }
            item.CREATE_DATE = System.DateTime.Now;
            logger.Info("TND_MAP_FP=" + item.ToString());
            return item;
        }
        #endregion
        #region 清單設備資料轉換 
        // * 處理清單設備，進入點
        public List<TND_MAP_DEVICE> ConvertDataForMapDEVICE(string projectId)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + ":工率");//因為與工率共用一個EXCEL SHEET
                sheet = (HSSFSheet)hssfworkbook.GetSheet("工率");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + ":工率");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("工率");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有設備清單資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[工率]資料");
            }
            return ConverData2MapDEVICE();
        }
        /**
         * 轉換圖算數量:設備清單
         * */
        protected List<TND_MAP_DEVICE> ConverData2MapDEVICE()
        {
            IRow row = null;
            List<TND_MAP_DEVICE> lstMapDEVICE = new List<TND_MAP_DEVICE>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..(表頭)
            while (iRowIndex < (4))
            {
                rows.MoveNext();
                iRowIndex++;
            }
            //循序處理每一筆資料之欄位!!
            iRowIndex++;
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                ICell cellPk = row.GetCell(0);
                //ICell CellQty = row.GetCell(11);
                //將各Row 資料寫入物件內
                //0.PK	11.設備數量 
                if (null == cellPk || cellPk.ToString().ToUpper() != "END")
                {
                    lstMapDEVICE.Add(convertRow2TndMapDEVICE(row, iRowIndex));
                }
                else
                {
                    logErrorMessage("Step1 ;取得圖算數量設備清單:" + lstMapDEVICE.Count + "筆");
                    logger.Info("Finish convert Job : count=" + lstMapDEVICE.Count);
                    return lstMapDEVICE;
                }
                iRowIndex++;
            }
            logger.Info("MAP_DEVICE Count:" + iRowIndex);
            return lstMapDEVICE;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件
         * */
        private TND_MAP_DEVICE convertRow2TndMapDEVICE(IRow row, int excelrow)
        {
            TND_MAP_DEVICE item = new TND_MAP_DEVICE();
            item.PROJECT_ID = projId;
            ICell cellPk = row.GetCell(0);
            ICell CellQty = row.GetCell(11);
            if (row.Cells[0].ToString().Trim() != "")//0.項次
            {
                item.PROJECT_ITEM_ID = row.Cells[0].ToString();
            }
            if (null != CellQty)//1.設備數量
            {
                try
                {
                    decimal dQty = decimal.Parse(CellQty.ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + dQty);
                    item.QTY = dQty;
                }
                catch (Exception e)
                {
                    logger.Error("data format Error on ExcelRow=" + excelrow + ",Cells[11].value=" + CellQty.ToString());
                    logger.Error(e);
                }
            }
            item.CREATE_DATE = System.DateTime.Now;
            logger.Info("TND_MAP_DEVICE=" + item.ToString());
            return item;
        }
        #endregion
        #region 給排水資料轉換 
        /**
         * 取得給排水Sheet 資料
         * */
        public List<TND_MAP_PLU> ConvertDataForMapPLU(string projectId)
        {
            projId = projectId;
            //1.依據檔案副檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
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
        protected List<TND_MAP_PLU> ConverData2MapPLU()
        {
            IRow row = null;
            List<TND_MAP_PLU> lstMapPLU = new List<TND_MAP_PLU>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..(表頭)
            while (iRowIndex < (1))
            {
                rows.MoveNext();
                iRowIndex++;
                //row = (IRow)rows.Current;
                //logger.Debug("skip data Excel Value:" + row.Cells[0].ToString() + "," + row.Cells[1] + "," + row.Cells[2]);
            }
            //循序處理每一筆資料之欄位!!
            iRowIndex++;
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                //將各Row 資料寫入物件內
                //0.項次	1.圖號	2.棟別	3.一次側位置	4.一次側名稱	5.二次側名稱	6.二次側位置	7.管材名稱	8管數/組	9管組數	10管長度/組數	11管總長
                if (row.Cells.Count != 0 && row.Cells[0].ToString().ToUpper() != "END")
                {
                    logger.Debug("Excel Value:" + row.Cells[0].ToString());
                    TND_MAP_PLU item = convertRow2TndMapPLU(row, iRowIndex);
                    if (null != item)
                    {
                        lstMapPLU.Add(item);
                    }
                }
                else
                {
                    logErrorMessage("Step1 ;取得圖算數量給排水:" + lstMapPLU.Count + "筆");
                    logger.Info("Finish convert Job : count=" + lstMapPLU.Count);
                    return lstMapPLU;
                }
                iRowIndex++;
            }
            logger.Info("MAP_PLU Count:" + iRowIndex);
            return lstMapPLU;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件
         * */
        private TND_MAP_PLU convertRow2TndMapPLU(IRow row, int excelrow)
        {
            TND_MAP_PLU item = new TND_MAP_PLU();
            item.PROJECT_ID = projId;
            logger.Debug("cell count:" + row.PhysicalNumberOfCells + ",Row:" + excelrow);
            try
            {
                #region parser rowdata about MAP_PLU
                if (row.Cells[0].ToString().Trim() != "")//0.項次
                {
                    item.EXCEL_ITEM = row.Cells[0].ToString();
                }
                if (row.Cells[1].ToString().Trim() != "")//1.圖號
                {
                    item.MAP_NO = row.Cells[1].ToString();
                }
                if (row.Cells[2].ToString().Trim() != "")//2.棟別
                {
                    item.BUILDING_NO = row.Cells[2].ToString();
                }

                if (row.Cells[3].ToString().Trim() != "")//3.一次側位置
                {
                    item.PRIMARY_SIDE = row.Cells[3].ToString();
                }
                if (row.Cells[4].ToString().Trim() != "")//4.一次側名稱
                {
                    item.PRIMARY_SIDE_NAME = row.Cells[4].ToString();
                }

                if (row.Cells[5].ToString().Trim() != "")//5.二次側名稱
                {
                    item.SECONDARY_SIDE = row.Cells[5].ToString();
                }
                if (row.Cells[6].ToString().Trim() != "")//6.二次側位置
                {
                    item.SECONDARY_SIDE_NAME = row.Cells[6].ToString();
                }

                if (row.Cells[7].ToString().Trim() != "")//7.管材名稱
                {
                    item.PIPE_NAME = row.Cells[7].ToString();
                }

                if (row.Cells[8].ToString().Trim() != "")// 8管數/組
                {
                    try
                    {
                        decimal dQty = decimal.Parse(row.Cells[8].ToString());
                        logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[8].ToString());
                        item.PIPE_COUNT_SET = dQty;
                    }
                    catch (Exception e)
                    {
                        logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[8].value=" + row.Cells[8].ToString());
                        logger.Error(e.Message);
                    }
                }

                if (row.Cells[9].ToString().Trim() != "")// 9.管組數	
                {
                    try
                    {
                        decimal dQty = decimal.Parse(row.Cells[9].ToString());
                        logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[9].ToString());
                        item.PIPE_SET_QTY = dQty;
                    }
                    catch (Exception e)
                    {
                        logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[9].value=" + row.Cells[9].ToString());
                        //   logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Item_Desc= " + projectItem.ITEM_DESC + ",value=" + row.Cells[3].ToString());
                        logger.Error(e.Message);
                    }
                }

                if (row.Cells[10].ToString().Trim() != "")// 10管長度/組數
                {
                    try
                    {
                        decimal dQty = decimal.Parse(row.Cells[10].ToString());
                        logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[10].ToString());
                        item.PIPE_LENGTH = dQty;
                    }
                    catch (Exception e)
                    {
                        logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[10].value=" + row.Cells[10].ToString());
                        logger.Error(e.Message);
                    }
                }


                if (row.Cells[11].ToString().Trim() != "")// 11管總長
                {
                    try
                    {
                        decimal dQty = decimal.Parse(row.Cells[11].ToString());
                        logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[11].ToString());
                        item.PIPE_TOTAL_LENGTH = dQty;
                    }
                    catch (Exception e)
                    {
                        logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[11].value=" + row.Cells[11].ToString());
                        logger.Error(e.Message);
                    }
                }

                item.CREATE_DATE = System.DateTime.Now;
                logger.Info("TND_MAP_PLU=" + item.ToString());
                return item;
                #endregion
            }
            catch (Exception ex)
            {
                logger.Debug("Row:" + excelrow + "cell count error :" + row.PhysicalNumberOfCells + ",ex=" + ex.StackTrace);
                return null;
            }
        }
        #endregion
        #region 弱電管線資料轉換 
        /**
         * 取得弱電管線Sheet 資料
         * */
        public List<TND_MAP_LCP> ConvertDataForMapLCP(string projectId)
        {
            projId = projectId;
            //1.依據檔案副檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + ":弱電管線");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("弱電管線");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + ":弱電管線");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("弱電管線");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有弱電管線資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[弱電管線]資料");
            }
            return ConverData2MapLCP();
        }
        /**
         * 轉換圖算數量:弱電管線
         * */
        protected List<TND_MAP_LCP> ConverData2MapLCP()
        {
            IRow row = null;
            List<TND_MAP_LCP> lstMapLCP = new List<TND_MAP_LCP>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..(表頭)
            while (iRowIndex < (1))
            {
                rows.MoveNext();
                iRowIndex++;
                // row = (IRow)rows.Current;
                // logger.Debug("skip data Excel Value:" + row.Cells[0].ToString() + "," + row.Cells[1] + "," + row.Cells[2]);
            }
            //循序處理每一筆資料之欄位!!
            iRowIndex++;
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                logger.Debug("Excel Row=" + iRowIndex + ",Excel Value:" + row.Cells[0].ToString());
                //將各Row 資料寫入物件內
                //0.項次	1.圖號	2.棟別	3.一次側位置	4.一次側名稱	5.二次側名稱	6.二次側位置	
                //7.線材名稱	8.條數/組	9.線組數	10.線長度/組數	11.線總長  12.地線名稱	13.地線條數	
                //14.地線總長  15.管材名稱1	--16.管長1	17.管組數1   18.管總長1	19.管材名稱2  20.管長2	21.管組數2  22.管總長2
                if (row.Cells.Count != 0 && row.Cells[0].ToString().ToUpper() != "END")
                {
                    logger.Debug("Excel Value:" + row.Cells[0].ToString() + row.Cells[1]);
                    lstMapLCP.Add(convertRow2TndMapLCP(row, iRowIndex));
                }
                else
                {
                    logErrorMessage("Step1 ;取得圖算數量弱電管線:" + lstMapLCP.Count + "筆");
                    logger.Info("Finish convert Job : count=" + lstMapLCP.Count);
                    return lstMapLCP;
                }
                iRowIndex++;
            }
            logger.Info("MAP_LCP Count:" + iRowIndex);
            return lstMapLCP;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件
         * */
        private TND_MAP_LCP convertRow2TndMapLCP(IRow row, int excelrow)
        {
            TND_MAP_LCP item = new TND_MAP_LCP();
            item.PROJECT_ID = projId;
            int iMax = row.LastCellNum;
            for (int iCol = 0; iCol < iMax; iCol++)
            {
                logger.Debug("Excel Row:" + excelrow + ",MaxCol=" + iCol);
                var cell = row.GetCell(iCol);
                if (null != cell)
                {
                    switch (iCol)
                    {
                        case 0://0項次
                            item.EXCEL_ITEM = cell.ToString();
                            break;
                        case 1://1圖號
                            item.MAP_NO = cell.ToString();
                            break;
                        case 2://2棟別
                            item.BUILDING_NO = cell.ToString();
                            break;
                        case 3://3一次側位置
                            item.PRIMARY_SIDE = cell.ToString();
                            break;
                        case 4://4一次側	
                            item.PRIMARY_SIDE_NAME = cell.ToString();
                            break;
                        case 5://5二次側
                            item.SECONDARY_SIDE = cell.ToString();
                            break;
                        case 6://6二次側位置
                            item.SECONDARY_SIDE_NAME = cell.ToString();
                            break;
                        case 7://7線徑
                            item.WIRE_NAME = cell.ToString();
                            break;
                        case 8://8數量
                            try
                            {
                                decimal dQty = decimal.Parse(cell.ToString());
                                logger.Info("excelrow=" + excelrow + ",value=" + cell.ToString());
                                item.WIRE_QTY_SET = dQty;
                            }
                            catch (Exception e)
                            {
                                logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[8].value=" + cell.ToString());
                                logger.Error(e.Message);
                            }
                            break;
                        case 9:// 9組數
                            try
                            {
                                decimal dQty = decimal.Parse(cell.ToString());
                                logger.Info("excelrow=" + excelrow + ",value=" + cell.ToString());
                                item.WIRE_SET_CNT = dQty;
                            }
                            catch (Exception e)
                            {
                                logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[9].value=" + cell.ToString());
                                //   logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Item_Desc= " + projectItem.ITEM_DESC + ",value=" + row.Cells[3].ToString());
                                logger.Error(e.Message + "," + e.StackTrace);
                            }
                            break;
                        case 10://10線長
                            try
                            {
                                decimal dQty = decimal.Parse(cell.ToString());
                                logger.Info("excelrow=" + excelrow + ",value=" + cell.ToString());
                                item.WIRE_LENGTH = dQty;
                            }
                            catch (Exception e)
                            {
                                logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[10].value=" + cell.ToString());
                                logger.Error(e.Message + "," + e.StackTrace);
                            }
                            break;
                        case 11://11線總長
                            try
                            {
                                decimal dQty = decimal.Parse(cell.ToString());
                                logger.Info("excelrow=" + excelrow + ",value=" + cell.ToString());
                                item.WIRE_TOTAL_LENGTH = dQty;
                            }
                            catch (Exception e)
                            {
                                logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[11].value=" + cell.ToString());
                                logger.Error(e.Message + "," + e.StackTrace);
                            }
                            break;
                        case 12://12導線管徑標稱
                            item.PIPE_1_NAME = cell.ToString();
                            break;
                        case 13://13管長
                            try
                            {
                                decimal dQty = decimal.Parse(cell.ToString());
                                logger.Info("excelrow=" + excelrow + ",value=" + cell.ToString());
                                item.PIPE_1_LENGTH = dQty;
                            }
                            catch (Exception e)
                            {
                                logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[13].value=" + cell.ToString());
                                logger.Error(e.Message + "," + e.StackTrace);
                            }
                            break;
                        case 14://14數量
                            try
                            {
                                decimal dQty = decimal.Parse(cell.ToString());
                                logger.Info("excelrow=" + excelrow + ",value=" + cell.ToString());
                                item.PIPE_1_SET = dQty;
                            }
                            catch (Exception e)
                            {
                                logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[14].value=" + cell.ToString());
                                logger.Error(e.Message + "," + e.StackTrace);
                            }
                            break;
                        case 15://15管總長
                            try
                            {
                                decimal dQty = decimal.Parse(cell.ToString());
                                logger.Info("excelrow=" + excelrow + ",value=" + cell.ToString());
                                item.PIPE_1_TOTAL_LEN = dQty;
                            }
                            catch (FormatException fe)
                            {
                                logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[15].value=" + cell.ToString());
                                logger.Error(fe.Message + "," + fe.StackTrace);
                            }
                            catch (Exception e)
                            {
                                logger.Error(e.Message + "," + e.StackTrace);
                            }
                            break;
                    }

                }
            }

            item.CREATE_DATE = System.DateTime.Now;
            logger.Info("TND_MAP_LCP=" + item.ToString());
            return item;
        }
        #endregion
        #region 電氣管線資料轉換 
        /**
         * 取得電氣管線Sheet 資料
         * */
        public List<TND_MAP_PEP> ConvertDataForMapPEP(string projectId)
        {
            projId = projectId;
            //1.依據檔案副檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
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
        protected List<TND_MAP_PEP> ConverData2MapPEP()
        {
            IRow row = null;
            List<TND_MAP_PEP> lstMapPEP = new List<TND_MAP_PEP>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..(表頭)
            while (iRowIndex < (1))
            {
                rows.MoveNext();
                iRowIndex++;
                //row = (IRow)rows.Current;
                //logger.Debug("skip data Excel Value:" + row.Cells[0].ToString() + "," + row.Cells[1] + "," + row.Cells[2]);
            }
            //循序處理每一筆資料之欄位!!
            iRowIndex++;
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                //將各Row 資料寫入物件內
                //0.項次	1.圖號	2.棟別	3.一次側位置	4.一次側名稱	5.二次側名稱	6.二次側位置	7.線材名稱	8.條數/組	9.線組數	10.線長度/條數	11.線總長  12.地線名稱	13.地線條數	14.地線總長  15.管材名稱	16.管長	17.管組數   18.管總長	
                if (row.Cells.Count != 0 && row.Cells[0].ToString().ToUpper() != "END")
                {
                    logger.Debug("Excel Value:" + row.Cells[0].ToString() + row.Cells[1] + row.Cells[2]);
                    lstMapPEP.Add(convertRow2TndMapPEP(row, iRowIndex));
                }
                else
                {
                    logErrorMessage("Step1 ;取得圖算數量電氣管線:" + lstMapPEP.Count + "筆");
                    logger.Info("Finish convert Job : count=" + lstMapPEP.Count);
                    return lstMapPEP;
                }
                iRowIndex++;
            }
            logger.Info("MAP_PEP Count:" + iRowIndex);
            return lstMapPEP;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件
         * */
        private TND_MAP_PEP convertRow2TndMapPEP(IRow row, int excelrow)
        {
            TND_MAP_PEP item = new TND_MAP_PEP();
            item.PROJECT_ID = projId;
            if (row.Cells[0].ToString().Trim() != "")//0.項次
            {
                item.EXCEL_ITEM = row.Cells[0].ToString();
            }
            if (row.Cells[1].ToString().Trim() != "")//1.圖號
            {
                item.MAP_NO = row.Cells[1].ToString();
            }
            if (row.Cells[2].ToString().Trim() != "")//2.棟別
            {
                item.BUILDING_NO = row.Cells[2].ToString();
            }

            if (row.Cells[3].ToString().Trim() != "")//3.一次側位置
            {
                item.PRIMARY_SIDE = row.Cells[3].ToString();
            }
            if (row.Cells[4].ToString().Trim() != "")//4.一次側名稱
            {
                item.PRIMARY_SIDE_NAME = row.Cells[4].ToString();
            }

            if (row.Cells[5].ToString().Trim() != "")//5.二次側名稱
            {
                item.SECONDARY_SIDE = row.Cells[5].ToString();
            }
            if (row.Cells[6].ToString().Trim() != "")//6.二次側位置
            {
                item.SECONDARY_SIDE_NAME = row.Cells[6].ToString();
            }

            if (row.Cells[7].ToString().Trim() != "")//7.線材名稱
            {
                item.WIRE_NAME = row.Cells[7].ToString();
            }

            if (row.Cells[8].ToString().Trim() != "")// 8.條數/組
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[8].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[8].ToString());
                    item.WIRE_QTY_SET = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[8].value=" + row.Cells[8].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[9].ToString().Trim() != "")// 9.線組數
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[9].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[9].ToString());
                    item.WIRE_SET_CNT = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[9].value=" + row.Cells[9].ToString());
                    //   logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Item_Desc= " + projectItem.ITEM_DESC + ",value=" + row.Cells[3].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[10].ToString().Trim() != "")// 10.線長度/組數   
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[10].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[10].ToString());
                    item.WIRE_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[10].value=" + row.Cells[10].ToString());
                    logger.Error(e.Message);
                }
            }


            if (row.Cells[11].ToString().Trim() != "")// 11.線總長  
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[11].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[11].ToString());
                    item.WIRE_TOTAL_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[11].value=" + row.Cells[11].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[12].ToString().Trim() != "") //12.地線名稱
            {
                item.GROUND_WIRE_NAME = row.Cells[12].ToString();
            }

            if (row.Cells[13].ToString().Trim() != "") // 	13.地線條數	
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[13].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[13].ToString());
                    item.GROUND_WIRE_QTY = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[13].value=" + row.Cells[13].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[14].ToString().Trim() != "")// 14.地線總長 		  	
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[14].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[14].ToString());
                    item.GROUND_WIRE_TOTAL_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[14].value=" + row.Cells[14].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[15].ToString().Trim() != "") //15.管材名稱
            {
                item.PIPE_NAME = row.Cells[15].ToString();
            }

            if (row.Cells[16].ToString().Trim() != "")//16.管長
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[16].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[16].ToString());
                    item.PIPE_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[16].value=" + row.Cells[16].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[17].ToString().Trim() != "")//17.管組數
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[17].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[17].ToString());
                    item.PIPE_SET = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[17].value=" + row.Cells[17].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[18].ToString().Trim() != "")//18.管總長
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[18].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[18].ToString());
                    item.PIPE_TOTAL_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[18].value=" + row.Cells[18].ToString());
                    logger.Error(e.Message);
                }
            }

            item.CREATE_DATE = System.DateTime.Now;
            logger.Info("TND_MAP_PEP=" + item.ToString());
            return item;
        }
        #endregion
        protected void logErrorMessage(string message)
        {
            if (errorMessage == null)
            {
                errorMessage = message;
            }
            else
            {
                errorMessage = errorMessage + "<br/>" + message;
            }
        }

    }
    #region 功率下載表格格式處理區段
    public class WageFormToExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string wageFile = ContextService.strUploadPath + "\\wage_form.xlsx";
        string outputPath = ContextService.strUploadPath;

        IWorkbook hssfworkbook;
        ISheet sheet = null;
        string fileformat = "xlsx";
        //存放工率資料
        public TND_PROJECT project = null;
        public List<PROJECT_ITEM_WITH_WAGE> projectItems = null;
        public string errorMessage = null;
        string projId = null;
        //建立工率下載表格
        public void exportExcel(TND_PROJECT project, List<PROJECT_ITEM_WITH_WAGE> projectItems)
        {
            //1.讀取功率表格檔案
            InitializeWorkbook(wageFile);
            sheet = (XSSFSheet)hssfworkbook.GetSheet("工率");

            //2.填入表頭資料
            logger.Debug("Table Head_1=" + sheet.GetRow(1).Cells[0].ToString());
            sheet.GetRow(1).Cells[1].SetCellValue(project.PROJECT_ID);//專案編號
            logger.Debug("Table Head_2=" + sheet.GetRow(2).Cells[0].ToString());
            sheet.GetRow(2).Cells[1].SetCellValue(project.PROJECT_NAME);//專案名稱

            //3.填入表單明細
            int idxRow = 4;
            foreach (PROJECT_ITEM_WITH_WAGE item in projectItems)
            {
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                //PK(PROJECT_ITEM_ID) 項次 名稱 單位 數量 單價 備註 九宮格 次九宮格 主系統 次系統
                row.CreateCell(0).SetCellValue(item.PROJECT_ITEM_ID);//PK(PROJECT_ITEM_ID)
                row.CreateCell(1).SetCellValue(item.ITEM_ID);//項次
                logger.Debug("Wage :ITEM DESC=" + item.ITEM_DESC);
                row.CreateCell(2).SetCellValue(item.ITEM_DESC);//項目說明
                row.CreateCell(3).SetCellValue(item.ITEM_UNIT);// 單位
                if (null != item.ITEM_QUANTITY && item.ITEM_QUANTITY.ToString().Trim() != "")
                {
                    row.CreateCell(4).SetCellValue(double.Parse(item.ITEM_QUANTITY.ToString())); //數量
                }
                if (null != item.ITEM_UNIT_PRICE && item.ITEM_UNIT_PRICE.ToString().Trim() != "")
                {
                    row.CreateCell(5).SetCellValue(double.Parse(item.ITEM_UNIT_PRICE.ToString())); //單價
                }
                row.CreateCell(6).SetCellValue(item.ITEM_REMARK);// 備註
                row.CreateCell(7).SetCellValue(item.TYPE_CODE_1);// 九宮格
                row.CreateCell(8).SetCellValue(item.TYPE_CODE_2);// 次九宮格
                row.CreateCell(9).SetCellValue(item.SYSTEM_MAIN);// 主系統
                row.CreateCell(10).SetCellValue(item.SYSTEM_SUB);// 次系統
                //建立空白欄位
                for (int iTmp = 11; iTmp < 20; iTmp++)
                {
                    row.CreateCell(iTmp);
                }
                if (item.RATIO != null)
                {
                    row.Cells[12].SetCellValue(Convert.ToDouble(item.RATIO));
                }

                idxRow++;
            }
            //4.令存新檔至專案所屬目錄
            var file = new FileStream(outputPath + "\\" + project.PROJECT_ID + "\\" + project.PROJECT_ID + "_工率.xlsx", FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
        }
        public WageFormToExcel()
        {
        }
        public void InitializeWorkbook(string path)
        {
            //read the wage file via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
            //book1.xls is an Excel-2007-generated file, so some new unknown BIFF records are added. 
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                logger.Info("Read Excel File:" + path); if (file.Name.EndsWith(".xls"))
                {
                    logger.Debug("process excel file for office 2003");
                    //fileformat = "xls";
                    hssfworkbook = new HSSFWorkbook(file);
                }
                else
                {
                    logger.Debug("process excel file for office 2007");
                    hssfworkbook = new XSSFWorkbook(file);
                }
                file.Close();
            }
        }
        #region 工率資料轉換 
        /**
         * 取得工率Sheet 資料
         * */
        public List<TND_WAGE> ConvertDataForWage(string projectId)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + ":工率");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("工率");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + ":工率");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("工率");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有工率資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[工率]資料");
            }
            return ConverData2Wage();
        }
        /**
         * 轉換工率資料檔:工率
         * */
        protected List<TND_WAGE> ConverData2Wage()
        {
            IRow row = null;
            List<TND_WAGE> lstWage = new List<TND_WAGE>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..(表頭)
            while (iRowIndex < (4))
            {
                rows.MoveNext();
                iRowIndex++;
                //row = (IRow)rows.Current;
                //logger.Debug("skip data Excel Value:" + row.Cells[0].ToString() + "," + row.Cells[1] + "," + row.Cells[2]);
            }
            //循序處理每一筆資料之欄位!!
            iRowIndex++;
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                int i = 0;
                string slog = "";
                for (i = 0; i < row.Cells.Count; i++)
                {
                    slog = slog + "," + row.Cells[i];

                }
                logger.Debug("Excel Value:" + slog);
                //將各Row 資料寫入物件內
                //0.PK	1.設備數量 2.工率
                if (row.Cells[0].ToString().ToUpper() != "END")
                {
                    lstWage.Add(convertRow2TndWage(row, iRowIndex));
                }
                else
                {
                    //logErrorMessage("Step1 ;取得工率資料:" + projectItems.Count + "筆");
                    //logger.Info("Finish convert Job : count=" + projectItems.Count);
                    return lstWage;
                }
                iRowIndex++;
            }
            logger.Info("MAP_WAGE Count:" + iRowIndex);
            return lstWage;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件
         * */
        private TND_WAGE convertRow2TndWage(IRow row, int excelrow)
        {
            TND_WAGE item = new TND_WAGE();
            item.PROJECT_ID = projId;
            if (row.Cells[0].ToString().Trim() != "")//0.項次
            {
                item.PROJECT_ITEM_ID = row.Cells[0].ToString();
            }
            if (null != row.Cells[row.Cells.Count - 1].ToString().Trim() || row.Cells[row.Cells.Count - 1].ToString().Trim() != "")//1.工率
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[row.Cells.Count - 1].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[row.Cells.Count - 1].ToString());
                    item.RATIO = dQty;
                }
                catch (Exception e)
                {
                    logger.Error("data format Error on ExcelRow=" + excelrow + ",Cells[12].value=" + row.Cells[row.Cells.Count - 1].ToString());
                    logger.Error(e);
                }
            }
            item.CREATE_DATE = System.DateTime.Now;
            logger.Info("TND_WAGE=" + item.ToString());
            return item;
        }

        #endregion
        private void logErrorMessage(string message)
        {
            if (errorMessage == null)
            {
                errorMessage = message;
            }
            else
            {
                errorMessage = errorMessage + "<br/>" + message;
            }
        }
    }
    #endregion
    #region 詢價單格式處理區段
    public class InquiryFormToExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string templateFile = ContextService.strUploadPath + "\\Inquiry_form_template.xlsx";
        public string outputPath = ContextService.strUploadPath;

        IWorkbook hssfworkbook;
        ISheet sheet = null;
        string fileformat = "xlsx";
        //存放供應商報價單資料
        public TND_PROJECT_FORM form = null;
        public List<TND_PROJECT_FORM_ITEM> formItems = null;

        // string fileformat = "xlsx";
        //建立詢價單樣板
        public string exportExcel(TND_PROJECT_FORM form, List<TND_PROJECT_FORM_ITEM> formItems, bool isTemp)
        {
            //1.讀取樣板檔案
            InitializeWorkbook(templateFile);
            sheet = (XSSFSheet)hssfworkbook.GetSheet("詢價單");
            //2.填入表頭資料
            Service.TnderProjectService ts = new TnderProjectService();
            TND_PROJECT p = ts.getProjectById(form.PROJECT_ID);
            logger.Debug("Template Head_1=" + sheet.GetRow(2).Cells[0].ToString());
            sheet.GetRow(2).Cells[1].SetCellValue(p.PROJECT_NAME);//專案名稱
            logger.Debug("Template Head_2=" + sheet.GetRow(3).Cells[0].ToString());
            sheet.GetRow(3).Cells[1].SetCellValue(form.FORM_NAME);//採購項目:
            logger.Debug("Template Head_3=" + sheet.GetRow(4).Cells[0].ToString());
            sheet.GetRow(4).Cells[1].SetCellValue(form.OWNER_NAME);//承辦人:
            logger.Debug("Template Head_4=" + sheet.GetRow(5).Cells[0].ToString());
            sheet.GetRow(5).Cells[1].SetCellValue(form.OWNER_TEL);//聯絡電話:
            logger.Debug("Template Head_5=" + sheet.GetRow(6).Cells[0].ToString());
            sheet.GetRow(6).Cells[1].SetCellValue(form.OWNER_EMAIL);//EMAIL:
            logger.Debug("Template Head_6=" + sheet.GetRow(7).Cells[0].ToString());
            sheet.GetRow(7).Cells[1].SetCellValue(form.OWNER_FAX);//FAX:
            logger.Debug("Template Head_1=" + sheet.GetRow(2).Cells[2].ToString());
            sheet.GetRow(2).Cells[3].SetCellValue(form.SUPPLIER_ID);//廠商名稱
            logger.Debug("Template Head_2=" + sheet.GetRow(3).Cells[2].ToString());
            sheet.GetRow(3).Cells[3].SetCellValue(form.CONTACT_NAME);//聯絡人
            logger.Debug("Template Head_3=" + sheet.GetRow(4).Cells[2].ToString());
            sheet.GetRow(4).Cells[3].SetCellValue(form.CONTACT_EMAIL);//電子信箱
            logger.Debug("Template Head_4=" + sheet.GetRow(5).Cells[2].ToString());
            sheet.GetRow(5).Cells[3].SetCellValue((form.DUEDATE).ToString());//報價期限
            logger.Debug("Template Head_5=" + sheet.GetRow(6).Cells[2].ToString());
            sheet.GetRow(6).Cells[3].SetCellValue(form.FORM_ID);//編號

            //3.填入表單明細
            int idxRow = 9;
            foreach (TND_PROJECT_FORM_ITEM item in formItems)
            {
                IRow row = sheet.GetRow(idxRow);
                //項次 項目說明    單位 數量  單價 複價  備註
                //row.Cells[0].SetCellValue(idxRow - 8);///項次
                row.Cells[0].SetCellValue(null == item.ITEM_ID ? "" : item.ITEM_ID);///項次
                logger.Debug("Inquiry :ITEM DESC=" + item.ITEM_DESC);
                row.Cells[1].SetCellValue(item.ITEM_DESC);//項目說明
                row.Cells[2].SetCellValue(item.ITEM_UNIT);// 單位
                if (null != item.ITEM_QTY && item.ITEM_QTY.ToString().Trim() != "")
                {
                    row.Cells[3].SetCellValue(double.Parse(item.ITEM_QTY.ToString())); //數量
                }
                //增加空白詢價單時，不寫入單價資料
                if (null != item.ITEM_UNIT_PRICE && item.ITEM_UNIT_PRICE.ToString().Trim() != "" && !isTemp)
                {
                    row.Cells[4].SetCellValue(double.Parse(item.ITEM_UNIT_PRICE.ToString())); //單價
                }
                // row.Cells[4].SetCellValue(idxRow - 8);//單價
                // row.Cells[5].SetCellValue(idxRow - 8);複價
                row.Cells[6].SetCellValue(item.ITEM_REMARK);// 備註
                                                            //建立空白欄位
                for (int iTmp = 7; iTmp < 27; iTmp++)
                {
                    row.CreateCell(iTmp);
                }
                //填入標單項次編號 PROJECT_ITEM_ID
                row.Cells[26].SetCellValue(item.PROJECT_ITEM_ID);
                idxRow++;
            }
            //4.另存新檔至專案所屬目錄 (增加Temp for zip 打包使用
            string fileLocation = null;
            if (isTemp)
            {
                fileLocation = outputPath + "\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\Temp\\" + form.FORM_NAME + "\\" + form.FORM_NAME + "_空白.xlsx";
            }
            else if (form.SUPPLIER_ID == null || form.SUPPLIER_ID == "")
            {
                fileLocation = outputPath + "\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\" + form.FORM_NAME + "_空白.xlsx";
            }
            else
            {
                fileLocation = outputPath + "\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\" + form.FORM_ID + ".xlsx";
            }
            var file = new FileStream(fileLocation, FileMode.Create);
            logger.Info("new file name =" + file.Name + ",path=" + file.Position);
            hssfworkbook.Write(file);
            file.Close();
            return fileLocation;
        }

        private void InitializeWorkbook(string path)
        {
            //read the template via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
            //book1.xls is an Excel-2007-generated file, so some new unknown BIFF records are added. 
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                logger.Info("Read Excel File:" + path); if (file.Name.EndsWith(".xls"))
                {
                    logger.Debug("process excel file for office 2003");
                    //fileformat = "xls";
                    hssfworkbook = new HSSFWorkbook(file);
                }
                else
                {
                    logger.Debug("process excel file for office 2007");
                    hssfworkbook = new XSSFWorkbook(file);
                }
                file.Close();
            }
        }
        public void convertInquiry2Project(string fileExcel, string projectid, string iswage)
        {
            //1.讀取供應商報價單\
            InitializeWorkbook(fileExcel);
            //2.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟廠商報價單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat);
                sheet = (HSSFSheet)hssfworkbook.GetSheet("詢價單");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat);
                sheet = (XSSFSheet)hssfworkbook.GetSheet("詢價單");
            }
            //3,讀取Sheet (預設詢價單，否則抓第一張)
            if (null == sheet)
            {
                sheet = (XSSFSheet)hssfworkbook.GetSheetAt(0);
            }
            //3.讀取檔頭 資料
            //專案名稱
            form = new TND_PROJECT_FORM();
            //專案名稱:	P0120
            logger.Debug(sheet.GetRow(2).Cells[0].ToString() + "," + sheet.GetRow(2).Cells[1]);
            form.PROJECT_ID = projectid;
            //工資報價單標記
            form.ISWAGE = iswage;
            //廠商名稱:	Supplier
            logger.Debug(sheet.GetRow(2).Cells[2].ToString() + "," + sheet.GetRow(2).Cells[3]);
            form.SUPPLIER_ID = sheet.GetRow(2).Cells[3].ToString(); //用供應商名稱暫代供應商編號
                                                                    //採購項目:	 詢價單名稱	
            logger.Debug(sheet.GetRow(3).Cells[0].ToString() + "," + sheet.GetRow(3).Cells[1]);
            form.FORM_NAME = sheet.GetRow(3).Cells[1].ToString();
            //聯絡人:	contact
            logger.Debug(sheet.GetRow(3).Cells[2].ToString() + "," + sheet.GetRow(3).Cells[3]);
            form.CONTACT_NAME = sheet.GetRow(3).Cells[3].ToString();
            //承辦人:
            logger.Debug(sheet.GetRow(4).Cells[0].ToString() + "," + sheet.GetRow(4).Cells[1]);
            form.OWNER_NAME = sheet.GetRow(4).Cells[1].ToString();
            //電子信箱:	contact@email.com
            logger.Debug(sheet.GetRow(4).Cells[2].ToString() + "," + sheet.GetRow(4).Cells[3]);
            form.CONTACT_EMAIL = sheet.GetRow(4).Cells[3].ToString();
            //聯絡電話:	08888888				
            logger.Debug(sheet.GetRow(5).Cells[0].ToString() + "," + sheet.GetRow(5).Cells[1]);
            form.OWNER_TEL = sheet.GetRow(5).Cells[1].ToString();
            //報價期限:	2017/1/25
            try
            {
                logger.Debug(sheet.GetRow(5).Cells[2].ToString() + "," + sheet.GetRow(5).Cells[3].ToString() + "," + sheet.GetRow(5).Cells[3].CellType);
                if (null == sheet.GetRow(5).Cells[3] || "" == sheet.GetRow(5).Cells[3].ToString())
                {
                    form.DUEDATE = DateTime.Now;
                }
                else
                {
                    string[] aryDate = sheet.GetRow(5).Cells[3].ToString().Split('/');
                    int intYear = int.Parse(aryDate[0]); ;
                    int intMonth = int.Parse(aryDate[1]);
                    int intDay = int.Parse(aryDate[2]);
                    if (intYear < 1900)
                    {
                        intYear = intYear + 1911;
                    }
                    DateTime dtDueDate = new DateTime(intYear, intMonth, intDay);
                    form.DUEDATE = dtDueDate;
                }
                logger.Debug("form.DUEDATE:" + form.DUEDATE);
            }
            catch (Exception ex)
            {
                logger.Error("Datetime format error: " + ex.Message);
                form.DUEDATE = DateTime.Now;
                // throw new Exception("日期格式有錯(YYYY/MM/DD");
            }
            //電子信箱:	admin@topmep
            logger.Debug(sheet.GetRow(6).Cells[0].ToString() + "," + sheet.GetRow(6).Cells[1]);
            form.OWNER_EMAIL = sheet.GetRow(6).Cells[1].ToString();
            //編號: REF - 001
            try
            {
                logger.Debug("REF_ID=" + sheet.GetRow(6).Cells[2].ToString() + "," + sheet.GetRow(6).Cells[3]);
                form.FORM_ID = sheet.GetRow(6).Cells[3].ToString().Trim();
            }
            catch (Exception ex)
            {
                form.FORM_ID = "";
                logger.Error("Not Reference ID:" + ex.Message);
            }

            //FAX:
            if (sheet.GetRow(7).GetCell(1) != null)
            {
                logger.Debug(sheet.GetRow(7).Cells[0].ToString());
                form.OWNER_FAX = sheet.GetRow(7).Cells[1].ToString();
            }
            else
            {
                form.OWNER_FAX = "";
            }

            //3.取得表單明細,逐行讀取資料
            IRow row = null;
            int iRowIndex = 9; //0 表 Row 1
            bool hasMore = true;
            //循序處理每一筆資料之欄位!!
            formItems = new List<TND_PROJECT_FORM_ITEM>();
            while (hasMore)
            {
                row = sheet.GetRow(iRowIndex);
                logger.Info("excel rowid=" + iRowIndex + ",phyCount=" + row.PhysicalNumberOfCells + ",cells count" + row.Cells.Count);
                if (null == row || row.Cells[0].ToString().ToUpper() == "END")
                {
                    return;
                }
                if (row.Cells.Count < 6)
                {
                    logger.Info("Row Index=" + iRowIndex + "column count has wrong" + row.Cells.Count);
                    throw new Exception("詢價單明細欄位有問題，請調整欄位相關資料(" + row.Cells.Count + ")");
                }
                else
                {
                    try
                    {
                        logger.Debug("row id=" + iRowIndex + "Cells Count=" + row.Cells.Count + ",form item vllue:" + row.Cells[0].ToString() + ","
                            + row.Cells[1] + "," + row.Cells[2] + "," + row.Cells[3] + "," + ","
                            + row.Cells[4] + "," + "," + row.Cells[5] + "," + row.Cells[6] + ",project item id=" + row.Cells[row.PhysicalNumberOfCells - 1]);
                        TND_PROJECT_FORM_ITEM item = new TND_PROJECT_FORM_ITEM();
                        item.ITEM_DESC = row.Cells[1].ToString();
                        item.ITEM_UNIT = row.Cells[2].ToString();
                        //標單數量
                        if (null != row.Cells[3] && row.Cells[3].ToString() != "")
                        {
                            decimal dQty = decimal.Parse(row.Cells[3].ToString());
                            item.ITEM_QTY = dQty;
                        }
                        else
                        {
                            logger.Warn("Qty format error:" + iRowIndex);
                        }
                        //報價單單價
                        if (null != row.Cells[4] && row.Cells[4].ToString() != "")
                        {
                            decimal dUnitPrice = decimal.Parse(row.Cells[4].ToString());
                            item.ITEM_UNIT_PRICE = dUnitPrice;
                        }
                        else
                        {
                            logger.Warn("UnitPrice format error:" + iRowIndex);
                        }

                        item.ITEM_REMARK = row.Cells[6].ToString();
                        logger.Info("Project ITEM ID=" + row.Cells[row.Cells.Count - 1].ToString());
                        item.PROJECT_ITEM_ID = row.Cells[row.Cells.Count - 1].ToString();
                        if (null != item.PROJECT_ITEM_ID && "" != item.PROJECT_ITEM_ID)
                        {
                            formItems.Add(item);
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error("Row Data Error:" + iRowIndex);
                        logger.Error(ex.GetType() + ":" + ex.StackTrace);
                    }
                }
                iRowIndex++;
            }
        }
    }
    #endregion
    #region 成本分析樣本輸入
    public class CostAnalysisOutput : ProjectItemFromExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string outputPath = ContextService.strUploadPath;
        string templateFile = ContextService.strUploadPath + "\\Cost_Analysis_Template_v1.xlsx";
        public string errorMsg = "";
        CostAnalysisDataService service = new CostAnalysisDataService();

        public TND_PROJECT project = null;
        public List<PROJECT_ITEM_WITH_WAGE> projectItems = null;

        XSSFCellStyle style = null;
        XSSFCellStyle styleNumber = null;

        public void exportExcel(string projectid)
        {
            //1 取得資料庫資料
            logger.Info("get data from DB,id=" + projectid);
            service.getProjectId(projectid);
            project = service.wageTable;
            projectItems = service.wageTableItem;

            //2.開啟檔案
            logger.Info("InitializeWorkbook");
            InitializeWorkbook(templateFile);
            style = ExcelStyle.getContentStyle(hssfworkbook);
            styleNumber = ExcelStyle.getNumberStyle(hssfworkbook);
            //3.期初成本 僅提供office 格式2007 
            getInitialQuotation();

            //4取得直接成本
            logger.Info("generate direct cost by projectid=" + project.PROJECT_ID);
            getDirectCostFromSouce(service.getDirectCost(project.PROJECT_ID));
            getDirectCostFromMap(service.DirectCost4Project);
            logger.Info("generate direct cost End!!");
            //5建立報價標單Sheet
            getFinalQuotation();
            //6建立報價標單Sheet
            logger.Info("generate sytem cost by projectid=" + project.PROJECT_ID);
            getSystemCost(service.getSystemCost(project.PROJECT_ID));
            logger.Info("generate sytem cost end!!");
            //7.建立預算參考表
            logger.Info("generate budget evaluate !!");
            getBudgetEva();
            logger.Info("generate budget evaluate end!!");
            //8.令存新檔至專案所屬目錄
            var file = new FileStream(outputPath + "\\" + project.PROJECT_ID + "\\" + project.PROJECT_ID + "_CostAnalysis.xlsx", FileMode.Create);
            logger.Info("output file=" + file.Name);
            hssfworkbook.Write(file);
            file.Close();
        }
        //寫入初期成本標單
        private void getInitialQuotation()
        {
            //2.寫入初期成本標單 僅提供office 格式2007 
            sheet = (XSSFSheet)hssfworkbook.GetSheet("期初成本");
            logger.Info("InitialQuotation  Head_1=" + sheet.GetRow(1).Cells[0].ToString() + "," + sheet.GetRow(1).Cells[1].ToString());
            sheet.GetRow(1).Cells[1].SetCellValue(project.PROJECT_ID);//專案編號
            sheet.GetRow(1).Cells[2].SetCellValue(project.PROJECT_NAME);//專案名稱
            sheet.GetRow(2).Cells[1].SetCellValue(project.LOCATION);//專案名稱

            int idxRow = 4;

            foreach (PROJECT_ITEM_WITH_WAGE item in projectItems)
            {
                logger.Debug("Row Id=" + idxRow + "," + item.ITEM_DESC);
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                //PK(PROJECT_ITEM_ID) 項次 名稱 單位 數量 單價 備註 九宮格 次九宮格 主系統 次系統
                row.CreateCell(0).SetCellValue(item.PROJECT_ITEM_ID);//PK(PROJECT_ITEM_ID)
                row.Cells[0].CellStyle = style;
                row.CreateCell(1).SetCellValue(item.ITEM_ID);//項次
                row.Cells[1].CellStyle = style;
                row.CreateCell(2).SetCellValue(item.ITEM_DESC);//項目說明
                row.Cells[2].CellStyle = style;
                row.CreateCell(3).SetCellValue(item.ITEM_UNIT);// 單位
                row.Cells[3].CellStyle = style;
                //標單數量
                if (null != item.ITEM_QUANTITY && item.ITEM_QUANTITY.ToString().Trim() != "")
                {
                    row.CreateCell(4).SetCellValue(double.Parse(item.ITEM_QUANTITY.ToString()));
                    row.Cells[4].CellStyle = style;
                }
                else
                {
                    row.CreateCell(4).SetCellValue("");
                }
                //圖算數量
                if (null != item.MAP_QTY && item.MAP_QTY.ToString().Trim() != "")
                {
                    row.CreateCell(5).SetCellValue(double.Parse(item.MAP_QTY.ToString()));
                    row.Cells[5].CellStyle = style;
                }
                else
                {
                    row.CreateCell(5).SetCellValue("");
                }
                //單價
                ICell cel6 = row.CreateCell(6);
                if (null != item.ITEM_UNIT_PRICE && item.ITEM_UNIT_PRICE.ToString().Trim() != "")
                {
                    logger.Debug("UNIT PRICE=" + item.ITEM_UNIT_PRICE);
                    cel6.SetCellValue(double.Parse(item.ITEM_UNIT_PRICE.ToString()));
                    cel6.CellStyle = styleNumber;
                }
                else
                {
                    cel6.SetCellValue("");
                    cel6.CellStyle = styleNumber;
                }
                ICell cel7 = row.CreateCell(7);
                if (null != item.ITEM_QUANTITY && null != item.ITEM_UNIT_PRICE)
                {
                    logger.Debug("Fomulor=" + "F" + (idxRow + 1) + "*G" + (idxRow + 1));
                    cel7.CellFormula = "F" + (idxRow + 1) + "*G" + (idxRow + 1);
                    cel7.CellStyle = styleNumber;
                }
                else
                {
                    cel7.SetCellValue("");
                    cel7.CellStyle = styleNumber;
                }
                row.CreateCell(8).SetCellValue((item.ITEM_REMARK == null ? "" : item.ITEM_REMARK));// 備註
                row.Cells[8].CellStyle = style;
                row.CreateCell(9).SetCellValue((item.TYPE_CODE_1 == null ? "" : item.TYPE_CODE_1));// 九宮格
                row.Cells[9].CellStyle = style;
                row.CreateCell(10).SetCellValue(item.TYPE_CODE_2 == null ? "" : item.TYPE_CODE_2);// 次九宮格
                row.Cells[10].CellStyle = style;
                row.CreateCell(11).SetCellValue((item.SYSTEM_MAIN == null ? "" : item.SYSTEM_MAIN));// 主系統
                row.Cells[11].CellStyle = style;
                row.CreateCell(12).SetCellValue((item.SYSTEM_SUB == null ? "" : item.SYSTEM_SUB));// 次系統
                                                                                                  //  row.Cells[12].CellStyle = style;
                if (null != item.RATIO)
                {
                    row.CreateCell(13).SetCellValue(double.Parse(item.RATIO.ToString()));// 工率
                    ICell cel14 = row.CreateCell(14);
                    cel14.CellFormula = "E" + (idxRow + 1) + "*N" + (idxRow + 1) + "*N3";
                }
                else
                {
                    row.CreateCell(13).SetCellValue("");// 工率
                    row.CreateCell(14).SetCellValue("");// 工資試算
                }
                row.Cells[13].CellStyle = style;
                row.Cells[14].CellStyle = style;
                idxRow++;
            }
            logger.Info("InitialQuotation finish!!");
            if (projectItems.Count == 0) { errorMsg = "標單資料不完整!!</br>"; }
        }
        //報價標單 
        private void getFinalQuotation()
        {
            //2.寫入報價標單 僅提供office 格式2007 
            sheet = (XSSFSheet)hssfworkbook.GetSheet("報價標單");
            logger.Debug("InitialQuotation  Head_1=" + sheet.GetRow(1).Cells[0].ToString() + "," + sheet.GetRow(1).Cells[1].ToString());
            sheet.GetRow(1).Cells[1].SetCellValue(project.PROJECT_ID);//專案編號
            sheet.GetRow(1).Cells[2].SetCellValue(project.PROJECT_NAME);//專案名稱
            sheet.GetRow(2).Cells[1].SetCellValue(project.LOCATION);//專案名稱

            int idxRow = 4;
            foreach (TND_PROJECT_ITEM item in projectItems)
            {
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                logger.Info("Row Id=" + idxRow);
                //PK(PROJECT_ITEM_ID) 項次 名稱 單位 數量 單價 備註 九宮格 次九宮格 主系統 次系統
                row.CreateCell(0).SetCellValue(item.PROJECT_ITEM_ID);//PK(PROJECT_ITEM_ID)
                row.CreateCell(1).SetCellValue(item.ITEM_ID);//項次
                row.CreateCell(2).SetCellValue(item.ITEM_DESC);//項目說明
                row.CreateCell(3).SetCellValue(item.ITEM_UNIT);// 單位
                row.CreateCell(4).SetCellValue("");
                if (null != item.ITEM_QUANTITY && item.ITEM_QUANTITY.ToString().Trim() != "")
                {
                    row.Cells[4].SetCellValue(double.Parse(item.ITEM_QUANTITY.ToString())); //數量
                }
                //單價先不填入
                row.CreateCell(5).SetCellValue("");
                if (null != item.ITEM_UNIT_PRICE && item.ITEM_UNIT_PRICE.ToString().Trim() != "")
                {
                    //row.CreateCell(5).SetCellValue(""); //單價
                }
                row.CreateCell(6).SetCellValue("");
                row.CreateCell(7).SetCellValue(item.ITEM_REMARK);// 備註
                row.CreateCell(8).SetCellValue(item.TYPE_CODE_1);// 九宮格
                row.CreateCell(9).SetCellValue(item.TYPE_CODE_2);// 次九宮格
                row.CreateCell(10).SetCellValue(item.SYSTEM_MAIN);// 主系統
                row.CreateCell(11).SetCellValue(item.SYSTEM_SUB);// 次系統
                logger.Debug("set cell style rowid=" + idxRow);
                foreach (ICell c in row.Cells)
                {
                    c.CellStyle = style;
                }
                idxRow++;
            }
        }
        //取得標單成本
        private void getDirectCostFromSouce(List<DirectCost> direcCostItems)
        {
            //2.寫入直接成本 僅提供office 格式2007 
            sheet = (XSSFSheet)hssfworkbook.GetSheet("標單成本");
            int idxRow = 1;
            foreach (DirectCost item in direcCostItems)
            {
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                logger.Info("Row Id=" + idxRow);
                //主九宮格編碼、次九宮格編碼、分項名稱(成本價)、材料報價、折扣率、材料成本、標單工資、複價、投標係數、投標複價、比例、項數
                //主九宮格編碼
                row.CreateCell(0).SetCellValue(item.MAINCODE);
                row.Cells[0].CellStyle = style;
                //次九宮格編碼
                if (null != item.SUB_CODE && item.SUB_CODE.ToString().Trim() != "")
                {
                    row.CreateCell(1).SetCellValue(double.Parse(item.SUB_CODE.ToString()));
                    row.Cells[1].CellStyle = style;
                }
                else
                {
                    row.CreateCell(1).SetCellValue("");
                    row.Cells[1].CellStyle = style;
                }
                //分項名稱
                logger.Debug("ITEM DESC=" + item.MAINCODE_DESC);
                row.CreateCell(2).SetCellValue(item.MAINCODE_DESC + "-" + item.SUB_DESC);
                row.Cells[2].CellStyle = style;
                //材料報價
                if (null != item.MATERIAL_COST && item.MATERIAL_COST.ToString().Trim() != "")
                {
                    row.CreateCell(3).SetCellValue(double.Parse(item.MATERIAL_COST.ToString()));
                    row.Cells[3].CellStyle = styleNumber;
                }
                else
                {
                    row.CreateCell(3).SetCellValue("");
                    row.Cells[3].CellStyle = styleNumber;
                }
                //折扣率 預設100%
                row.CreateCell(4).SetCellValue("100%");
                row.Cells[4].CellStyle = style;
                //材料成本
                row.CreateCell(5).CellFormula = "D" + (idxRow + 1) + "*E" + (idxRow + 1);
                row.Cells[5].CellStyle = styleNumber;
                //標單工資from 工率
                if (null != item.MAN_DAY && item.MAN_DAY.ToString().Trim() != "")
                {
                    double manday = double.Parse(item.MAN_DAY.ToString());
                    row.CreateCell(6).SetCellFormula(manday + "*期初成本!N3");
                    row.Cells[6].CellStyle = styleNumber;
                }
                else
                {
                    row.CreateCell(6).SetCellValue("");
                    row.Cells[6].CellStyle = style;
                }
                //複價//報價複價(填入公式)
                row.CreateCell(7).CellFormula = "F" + (idxRow + 1) + "+G" + (idxRow + 1);
                row.Cells[7].CellStyle = styleNumber;
                //投標係數、投標複價、比例、項數
                row.CreateCell(8).SetCellValue("100%");// 投標係數
                row.Cells[8].CellStyle = style;
                // 投標複價
                row.CreateCell(9).CellFormula = "F" + (idxRow + 1) + "*I" + (idxRow + 1); ;
                row.Cells[9].CellStyle = styleNumber;
                //比例
                row.CreateCell(10).SetCellValue("0");// 比例
                row.Cells[10].CellStyle = style;
                if (null != item.ITEM_COUNT && item.ITEM_COUNT.ToString().Trim() != "")
                {
                    row.CreateCell(11).SetCellValue(int.Parse(item.ITEM_COUNT.ToString())); //項數
                    row.Cells[11].CellStyle = style;
                }
                idxRow++;
            }
            //錯誤控制
            if (direcCostItems.Count == 0)
            {
                errorMsg = errorMsg + "成本資料不完整,無法產生成本分析資料!!</br>";
                return;
            }
            //加入小計欄位
            logger.Debug("add row for summary:" + idxRow);
            IRow summaryRow = sheet.CreateRow(idxRow);//.GetRow(idxRow);
            for (int i = 0; i < 12; i++)
            {
                summaryRow.CreateCell(i);
            }
            summaryRow.Cells[2].SetCellValue("小計");
            summaryRow.Cells[3].SetCellFormula("SUM(D2:D" + (idxRow) + ")");
            summaryRow.Cells[3].CellStyle = styleNumber;
            summaryRow.Cells[5].SetCellFormula("SUM(F2:F" + (idxRow) + ")");
            summaryRow.Cells[5].CellStyle = styleNumber;
            summaryRow.Cells[6].SetCellFormula("SUM(G2:G" + (idxRow) + ")");
            summaryRow.Cells[6].CellStyle = styleNumber;
            summaryRow.Cells[7].SetCellFormula("SUM(H2:H" + (idxRow) + ")");
            summaryRow.Cells[7].CellStyle = styleNumber;
            summaryRow.Cells[9].SetCellFormula("SUM(J2:J" + (idxRow) + ")");
            summaryRow.Cells[9].CellStyle = styleNumber;
            summaryRow.Cells[11].SetCellFormula("SUM(L2:L" + (idxRow) + ")");
            summaryRow.Cells[11].CellStyle = styleNumber;
            logger.Debug("set cell style rowid=" + idxRow);
        }
        //產生圖算成本
        private void getDirectCostFromMap(List<DirectCost> direcCostItems)
        {
            //2.寫入直接成本 僅提供office 格式2007 
            sheet = (XSSFSheet)hssfworkbook.GetSheet("圖算成本");
            int idxRow = 1;
            foreach (DirectCost item in direcCostItems)
            {
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                logger.Info("Row Id=" + idxRow);
                //主九宮格編碼 次九宮格編碼  分項名稱 材料報價    折扣率 材料成本    圖算工資 複價  項數
                //主九宮格編碼
                row.CreateCell(0).SetCellValue(item.MAINCODE);
                row.Cells[0].CellStyle = style;
                //次九宮格編碼
                if (null != item.SUB_CODE && item.SUB_CODE.ToString().Trim() != "")
                {
                    row.CreateCell(1).SetCellValue(double.Parse(item.SUB_CODE.ToString()));
                }
                else
                {
                    row.CreateCell(1).SetCellValue("");
                }
                row.Cells[1].CellStyle = style;
                //分項名稱
                logger.Debug("ITEM DESC=" + item.MAINCODE_DESC);
                row.CreateCell(2).SetCellValue(item.MAINCODE_DESC + "-" + item.SUB_DESC);
                row.Cells[2].CellStyle = style;
                //材料報價
                if (null != item.MATERIAL_COST_INMAP && item.MATERIAL_COST_INMAP.ToString().Trim() != "")
                {
                    row.CreateCell(3).SetCellValue(double.Parse(item.MATERIAL_COST_INMAP.ToString()));
                }
                else
                {
                    row.CreateCell(3).SetCellValue("");
                }
                row.Cells[3].CellStyle = styleNumber;
                //折扣率 預設100%
                row.CreateCell(4).SetCellValue("100%");
                row.Cells[4].CellStyle = style;
                //材料成本
                row.CreateCell(5).CellFormula = "D" + (idxRow + 1) + "*E" + (idxRow + 1);
                row.Cells[5].CellStyle = styleNumber;
                //圖算工資
                if (null != item.MAN_DAY_INMAP && item.MAN_DAY_INMAP.ToString().Trim() != "")
                {
                    double manday = double.Parse(item.MAN_DAY_INMAP.ToString());
                    row.CreateCell(6).SetCellFormula(manday + "*期初成本!N3");
                }
                else
                {
                    row.CreateCell(6).SetCellValue("");
                }
                row.Cells[6].CellStyle = styleNumber;
                //複價//報價複價(填入公式)
                row.CreateCell(7).CellFormula = "F" + (idxRow + 1) + "+G" + (idxRow + 1);
                row.Cells[7].CellStyle = styleNumber;
                //項數
                if (null != item.ITEM_COUNT && item.ITEM_COUNT.ToString().Trim() != "")
                {
                    row.CreateCell(8).SetCellValue(int.Parse(item.ITEM_COUNT.ToString())); //項數
                }
                idxRow++;
            }
            //錯誤控制
            if (direcCostItems.Count == 0)
            {
                errorMsg = errorMsg + "成本資料不完整,無法產生圖算成本資料!!</br>";
                return;
            }
            //加入小計欄位
            logger.Debug("add row for summary:" + idxRow);
            IRow summaryRow = sheet.CreateRow(idxRow);//.GetRow(idxRow);
            for (int i = 0; i < 9; i++)
            {
                summaryRow.CreateCell(i);
            }
            summaryRow.Cells[2].SetCellValue("小計");
            summaryRow.Cells[3].SetCellFormula("SUM(D2:D" + (idxRow) + ")");

            summaryRow.Cells[5].SetCellFormula("SUM(F2:F" + (idxRow) + ")");
            //加入工資單價
            summaryRow.Cells[6].SetCellFormula("SUM(G2:G" + (idxRow) + ")");
            summaryRow.Cells[7].SetCellFormula("SUM(H2:H" + (idxRow) + ")");
            summaryRow.Cells[8].SetCellFormula("SUM(I2:I" + (idxRow) + ")");
            logger.Debug("set cell style rowid=" + idxRow);
            foreach (ICell c in summaryRow.Cells)
            {
                c.CellStyle = style;
            }
        }
        //產生系統成本
        private void getSystemCost(List<SystemCost> systemCostItems)
        {
            //1.寫系統分析成本 僅提供office 格式2007 
            sheet = (XSSFSheet)hssfworkbook.GetSheet("系統分析");

            int idxRow = 1;
            foreach (SystemCost item in systemCostItems)
            {
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                logger.Info("Row Id=" + idxRow);
                //分項名稱(成本價)-主系統	分項名稱(成本價)-次系統	標單材料成本	圖算材料成本	標單工資	圖算工資	標單成本	圖算成本	項數
                row.CreateCell(0).SetCellValue(item.SYSTEM_MAIN);
                row.Cells[0].CellStyle = style;
                //分項名稱(成本價) - 次系統
                row.CreateCell(1).SetCellValue(item.SYSTEM_SUB);
                row.Cells[1].CellStyle = style;
                //標單材料成本//
                if (null != item.MATERIAL_COST && item.MATERIAL_COST.ToString().Trim() != "")
                {
                    row.CreateCell(2).SetCellValue(double.Parse(item.MATERIAL_COST.ToString()));
                }
                else
                {
                    row.CreateCell(2).SetCellValue("");
                }
                row.Cells[2].CellStyle = styleNumber;
                //圖算材料成本//
                if (null != item.MATERIAL_COST_INMAP && item.MATERIAL_COST_INMAP.ToString().Trim() != "")
                {
                    row.CreateCell(3).SetCellValue(double.Parse(item.MATERIAL_COST.ToString()));
                }
                else
                {
                    row.CreateCell(3).SetCellValue("");
                }
                row.Cells[3].CellStyle = styleNumber;
                //標單工資//
                if (null != item.MAN_DAY && item.MAN_DAY.ToString().Trim() != "")
                {
                    double manday = double.Parse(item.MAN_DAY.ToString());
                    row.CreateCell(4).SetCellFormula(manday + "*期初成本!N3");
                }
                else
                {
                    row.CreateCell(4).SetCellValue("");
                }
                row.Cells[4].CellStyle = styleNumber;
                //圖算工資
                if (null != item.MAN_DAY_INMAP && item.MAN_DAY_INMAP.ToString().Trim() != "")
                {
                    double manday = double.Parse(item.MAN_DAY_INMAP.ToString());
                    row.CreateCell(5).SetCellFormula(manday + "*期初成本!N3");
                }
                else
                {
                    row.CreateCell(5).SetCellValue("");
                }
                row.Cells[5].CellStyle = styleNumber;

                //標單成本               row.CreateCell(5).CellFormula = "D" + (idxRow + 1) + "*E" + (idxRow + 1);
                row.CreateCell(6).CellFormula = "C" + (idxRow + 1) + "+E" + (idxRow + 1);
                row.Cells[6].CellStyle = styleNumber;
                //圖算成本
                row.CreateCell(7).CellFormula = "D" + (idxRow + 1) + "+F" + (idxRow + 1);
                row.Cells[7].CellStyle = styleNumber;
                //項數
                if (null != item.ITEM_COUNT && item.ITEM_COUNT.ToString().Trim() != "")
                {
                    row.CreateCell(8).SetCellValue(int.Parse(item.ITEM_COUNT.ToString())); //項數
                }
                idxRow++;
            }
            //錯誤控制
            if (systemCostItems.Count == 0)
            {
                errorMsg = errorMsg + "成本資料不完整,無法產生系統成本資料!!</br>";
                return;
            }
            //加入小計欄位
            logger.Debug("add row for summary:" + idxRow);
            IRow summaryRow = sheet.CreateRow(idxRow);//.GetRow(idxRow);
            for (int i = 0; i < 9; i++)
            {
                summaryRow.CreateCell(i);
            }
            summaryRow.Cells[1].SetCellValue("直接工程費(小計)");
            summaryRow.Cells[2].SetCellFormula("SUM(C2:C" + (idxRow) + ")");
            summaryRow.Cells[2].CellStyle = styleNumber;
            summaryRow.Cells[3].SetCellFormula("SUM(D2:D" + (idxRow) + ")");
            summaryRow.Cells[3].CellStyle = styleNumber;
            summaryRow.Cells[4].SetCellFormula("SUM(E2:E" + (idxRow) + ")");
            summaryRow.Cells[4].CellStyle = styleNumber;
            summaryRow.Cells[5].SetCellFormula("SUM(F2:F" + (idxRow) + ")");
            summaryRow.Cells[5].CellStyle = styleNumber;
            summaryRow.Cells[6].SetCellFormula("SUM(G2:G" + (idxRow) + ")");
            summaryRow.Cells[6].CellStyle = styleNumber;
            summaryRow.Cells[7].SetCellFormula("SUM(H2:H" + (idxRow) + ")");
            summaryRow.Cells[7].CellStyle = styleNumber;
            summaryRow.Cells[8].SetCellFormula("SUM(I2:I" + (idxRow) + ")");
        }
        //預算參考表
        private void getBudgetEva()
        {
            List<DirectCost> typecodeItems = service.getDirectCost4BudgetRef(project.PROJECT_ID);
            //1.讀取預算表格檔案
            sheet = (XSSFSheet)hssfworkbook.GetSheet("預算參考表");
            //2.填入表頭資料
            logger.Debug("Table Head_1=" + sheet.GetRow(1).Cells[0].ToString());
            sheet.GetRow(1).Cells[1].SetCellValue(project.PROJECT_ID);//專案編號
            logger.Debug("Table Head_2=" + sheet.GetRow(2).Cells[0].ToString());
            sheet.GetRow(2).Cells[1].SetCellValue(project.PROJECT_NAME);//專案名稱
            //3.填入資料
            int idxRow = 4;
            foreach (DirectCost item in typecodeItems)
            {
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                logger.Debug("Row Id=" + idxRow + ",ITEM_DESC=" + item.MAINCODE_DESC);
                //主九宮格編碼、次九宮格編碼、主系統、次系統、分項名稱(成本價)、合約金額、材料成本、預算折扣率、預算金額
                //主九宮格編碼
                row.CreateCell(0).SetCellValue(item.MAINCODE);
                row.Cells[0].CellStyle = style;
                //次九宮格編碼
                row.CreateCell(1).SetCellValue("");
                if (null != item.SUB_CODE && item.SUB_CODE.ToString().Trim() != "")
                {
                    row.Cells[1].SetCellValue(double.Parse(item.SUB_CODE.ToString()));
                }
                row.Cells[1].CellStyle = style;
                //分項名稱
                row.CreateCell(2).SetCellValue(item.MAINCODE_DESC + "-" + item.SUB_DESC);
                row.Cells[2].CellStyle = style;
                //合約金額
                row.CreateCell(3).SetCellValue("--");
                if (null != item.CONTRACT_PRICE && item.CONTRACT_PRICE.ToString().Trim() != "")
                {
                    // row.Cells[3].SetCellValue(double.Parse(item.CONTRACT_PRICE.ToString()));
                    //  row.Cells[3].CellStyle = styleNumber;
                }
                //材料成本 與預算折扣率
                row.CreateCell(4).SetCellValue("");
                if (null != item.MATERIAL_COST_INMAP && item.MATERIAL_COST_INMAP.ToString().Trim() != "")
                {
                    row.Cells[4].SetCellValue(double.Parse(item.MATERIAL_COST_INMAP.ToString()));
                    row.Cells[4].CellStyle = styleNumber;
                }
                row.CreateCell(5).SetCellValue("100");
                row.Cells[5].CellStyle = style;
                //圖算*工率
                row.CreateCell(6).SetCellValue("");  //圖算*工率
                row.CreateCell(7).SetCellValue("");
                if (null != item.MAN_DAY_INMAP && item.MAN_DAY_INMAP.ToString().Trim() != "")
                {
                    row.Cells[6].SetCellFormula(item.MAN_DAY_INMAP.ToString() + "*G3");
                    row.Cells[7].SetCellValue("100");
                }
                row.Cells[6].CellStyle = styleNumber;
                row.Cells[7].CellStyle = style;
                //預算金額
                ICell cell8 = row.CreateCell(8);
                cell8.CellFormula = "(E" + (idxRow + 1) + "*F" + (idxRow + 1) + "/100)+(G" + (idxRow + 1) + "*H" + (idxRow + 1) + "/100)";
                cell8.CellStyle = styleNumber;
                logger.Debug("getBudget cell style rowid=" + idxRow);
                idxRow++;
            }
            //加入加總欄位
            IRow summaryRow = sheet.CreateRow(idxRow);
            summaryRow.CreateCell(0).SetCellValue("");
            summaryRow.Cells[0].CellStyle = style;
            summaryRow.CreateCell(1).SetCellValue("");
            summaryRow.Cells[1].CellStyle = style;
            summaryRow.CreateCell(2).SetCellValue("加總");
            summaryRow.Cells[2].CellStyle = style;
            summaryRow.CreateCell(3);
            summaryRow.Cells[3].SetCellFormula("SUM(D5:D" + idxRow + ")");
            summaryRow.Cells[3].CellStyle = styleNumber;
            summaryRow.CreateCell(4);
            summaryRow.Cells[4].SetCellFormula("SUM(E5:E" + idxRow + ")");
            summaryRow.Cells[4].CellStyle = styleNumber;
            summaryRow.CreateCell(5).SetCellValue("");
            summaryRow.Cells[5].CellStyle = style;
            summaryRow.CreateCell(6);
            summaryRow.Cells[6].SetCellFormula("SUM(G5:G" + idxRow + ")");
            summaryRow.Cells[6].CellStyle = styleNumber;
            summaryRow.CreateCell(7).SetCellValue("");
            summaryRow.Cells[7].CellStyle = style;
            summaryRow.CreateCell(8);
            summaryRow.Cells[8].SetCellFormula("SUM(I5:I" + idxRow + ")");
            summaryRow.Cells[8].CellStyle = styleNumber;
        }
    }
    public class ExcelStyle
    {
        //顯示有逗號區分的數值資料
        //HSSFCellStyle styleNumeric = (HSSFCellStyle)workbook.CreateCellStyle();
        //HSSFDataFormat format = (HSSFDataFormat)workbook.CreateDataFormat();
        //styleNumeric.DataFormat =format.GetFormat("###,##0");
        //cell.CellStyle = styleNumeric;
        public static XSSFCellStyle getNumberStyle(IWorkbook workbook)
        {
            XSSFCellStyle styleNumeric = (XSSFCellStyle)workbook.CreateCellStyle();
            styleNumeric.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            styleNumeric.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;//細實線
            styleNumeric.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;//虛線
            styleNumeric.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;//...  
            XSSFDataFormat format = (XSSFDataFormat)workbook.CreateDataFormat();
            styleNumeric.DataFormat = format.GetFormat("#,##0");
            return styleNumeric;
        }
        public static XSSFCellStyle getContentStyle(IWorkbook workbook)
        {
            XSSFCellStyle oStyle = (XSSFCellStyle)workbook.CreateCellStyle();

            //設定上下左右的框線
            oStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;//粗
            oStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;//細實線
            oStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;//虛線
            oStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;//...  
            return oStyle;
        }
        //顯示百分比符號
        //HSSFCellStyle stylePercent = (HSSFCellStyle)workbook.CreateCellStyle();
        //HSSFDataFormat format = (HSSFDataFormat)workbook.CreateDataFormat();
        //        stylePercent.DataFormat = format.GetFormat("#0.00%");
        //cell.CellStyle = stylePercent;
        //顯示日期格式
        //HSSFCellStyle styleDate = (HSSFCellStyle)workbook.CreateCellStyle();
        //HSSFDataFormat format = (HSSFDataFormat)workbook.CreateDataFormat();
        //        styleDate.DataFormat = format.GetFormat("yyyy-mm-dd");
        //cell.CellStyle = styleDate;

    }
    public class ProjectItem2Excel : ProjectItemFromExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string outputPath = ContextService.strUploadPath;
        string templateFile = ContextService.strUploadPath + "\\ProjectItem_Template_v1.xlsx";
        public string errorMsg = "";
        public TND_PROJECT project = null;
        public List<PROJECT_ITEM_WITH_WAGE> projectItems = null;
        CostAnalysisDataService service = new CostAnalysisDataService();

        public XSSFCellStyle style = null;
        public XSSFCellStyle styleNumber = null;

        public void exportExcel(string projectid)
        {
            //1 取得資料庫資料
            logger.Info("get data from DB,id=" + projectid);
            service.getProjectId(projectid);
            project = service.wageTable;
            projectItems = service.wageTableItem;

            //2.開啟檔案
            logger.Info("InitializeWorkbook");
            InitializeWorkbook(templateFile);
            style = ExcelStyle.getContentStyle(hssfworkbook);
            styleNumber = ExcelStyle.getNumberStyle(hssfworkbook);
            //3.標單品項 僅提供office 格式2007 
            getProjectItem();

            //4.令存新檔至專案所屬目錄
            var file = new FileStream(outputPath + "\\" + project.PROJECT_ID + "\\" + project.PROJECT_ID + "_標單明細.xlsx", FileMode.Create);
            logger.Info("output file=" + file.Name);
            hssfworkbook.Write(file);
            file.Close();
        }
        //寫入標單品項
        protected void getProjectItem()
        {
            //2.寫入初期成本標單 僅提供office 格式2007 
            sheet = (XSSFSheet)hssfworkbook.GetSheet("標單品項");
            logger.Debug("InitialQuotation  Head_1=" + sheet.GetRow(1).Cells[0].ToString() + "," + sheet.GetRow(1).Cells[1].ToString());
            sheet.GetRow(1).Cells[1].SetCellValue(project.PROJECT_ID);//專案編號
            sheet.GetRow(1).Cells[2].SetCellValue(project.PROJECT_NAME);//專案名稱
            sheet.GetRow(2).Cells[1].SetCellValue(project.LOCATION);//專案名稱

            int idxRow = 4;
            foreach (PROJECT_ITEM_WITH_WAGE item in projectItems)
            {
                logger.Debug("Row Id=" + idxRow);
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                //PK(PROJECT_ITEM_ID) 項次 名稱 單位 數量 單價 備註 九宮格 次九宮格 主系統 次系統,Excel 排序
                row.CreateCell(0).SetCellValue(item.PROJECT_ITEM_ID);//PK(PROJECT_ITEM_ID)
                row.Cells[0].CellStyle = style;
                row.CreateCell(1).SetCellValue(item.ITEM_ID);//項次
                row.Cells[1].CellStyle = style;
                logger.Debug("ITEM DESC=" + item.ITEM_DESC);
                row.CreateCell(2).SetCellValue(item.ITEM_DESC);//項目說明
                row.Cells[2].CellStyle = style;
                row.CreateCell(3).SetCellValue(item.ITEM_UNIT);// 單位
                row.Cells[3].CellStyle = style;
                //標單數量
                if (null != item.ITEM_QUANTITY && item.ITEM_QUANTITY.ToString().Trim() != "")
                {
                    row.CreateCell(4).SetCellValue(double.Parse(item.ITEM_QUANTITY.ToString()));
                    row.Cells[4].CellStyle = style;
                }
                else
                {
                    row.CreateCell(4).SetCellValue("");
                }
                //圖算數量
                if (null != item.MAP_QTY && item.MAP_QTY.ToString().Trim() != "")
                {
                    row.CreateCell(5).SetCellValue(double.Parse(item.MAP_QTY.ToString()));
                    row.Cells[5].CellStyle = style;
                }
                else
                {
                    row.CreateCell(5).SetCellValue("");
                }
                //單價
                ICell cel6 = row.CreateCell(6);
                if (null != item.ITEM_UNIT_PRICE && item.ITEM_UNIT_PRICE.ToString().Trim() != "")
                {
                    logger.Debug("UNIT PRICE=" + item.ITEM_UNIT_PRICE);
                    cel6.SetCellValue(double.Parse(item.ITEM_UNIT_PRICE.ToString()));
                    cel6.CellStyle = styleNumber;
                }
                else
                {
                    cel6.SetCellValue("");
                    cel6.CellStyle = styleNumber;
                }
                ICell cel7 = row.CreateCell(7);
                if (null != item.ITEM_QUANTITY && null != item.ITEM_UNIT_PRICE)
                {
                    logger.Debug("Fomulor=" + "F" + (idxRow + 1) + "*G" + (idxRow + 1));
                    cel7.CellFormula = "F" + (idxRow + 1) + "*G" + (idxRow + 1);
                    cel7.CellStyle = styleNumber;
                }
                else
                {
                    cel7.SetCellValue("");
                    cel7.CellStyle = styleNumber;
                }
                row.CreateCell(8).SetCellValue((item.ITEM_REMARK == null ? "" : item.ITEM_REMARK));// 備註
                row.Cells[8].CellStyle = style;
                row.CreateCell(9).SetCellValue((item.TYPE_CODE_1 == null ? "" : item.TYPE_CODE_1));// 九宮格
                row.Cells[9].CellStyle = style;
                row.CreateCell(10).SetCellValue(item.TYPE_CODE_2 == null ? "" : item.TYPE_CODE_2);// 次九宮格
                row.Cells[10].CellStyle = style;
                row.CreateCell(11).SetCellValue((item.SYSTEM_MAIN == null ? "" : item.SYSTEM_MAIN));// 主系統
                row.Cells[11].CellStyle = style;
                row.CreateCell(12).SetCellValue((item.SYSTEM_SUB == null ? "" : item.SYSTEM_SUB));// 次系統
                row.Cells[12].CellStyle = style;
                row.CreateCell(13).SetCellValue((item.EXCEL_ROW_ID == null ? "" : item.EXCEL_ROW_ID.ToString()));// Excel 排序
                row.Cells[13].CellStyle = style;
                row.CreateCell(14).SetCellValue((item.IN_CONTRACT == null ? "Y" : item.IN_CONTRACT.ToString()));// 合約內品項標記
                row.Cells[14].CellStyle = style;
                idxRow++;
            }
        }
        public void ConvertDataForTenderProject(string projectId)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId);
                sheet = (HSSFSheet)hssfworkbook.GetSheet("標單品項");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId);
                sheet = (XSSFSheet)hssfworkbook.GetSheet("標單品項");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有標單品項(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有標單品項(Sheet)");
            }
            ConvertExcelToTndProjectItem(5);
        }
        //轉換標單內容物件
        public new void ConvertExcelToTndProjectItem(int startrow)
        {
            IRow row = null;
            lstProjectItem = new List<TND_PROJECT_ITEM>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..
            while (iRowIndex < (startrow - 1))
            {
                rows.MoveNext();
                iRowIndex++;
                row = (IRow)rows.Current;
                logger.Debug("skip data Excel Value:" + row.Cells[0].ToString() + "," + row.Cells[1] + "," + row.Cells[2]);
            }
            //循序處理每一筆資料之欄位!!
            iRowIndex++;
            int itemId = 1;
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                logger.Debug("Cells Count=" + row.Cells.Count + ",Excel Value:" + row.Cells[0].ToString());
                //將各Row 資料寫入物件內
                //項次,名稱,單位,數量,單價,複價,備註,九宮格,次九宮格,主系統,次系統
                if (row.Cells[0].ToString().ToUpper() != "END")
                {
                    TND_PROJECT_ITEM projectItem = convertRow2TndProjectItem4Refresh(itemId, row, iRowIndex);
                    if (null != projectItem)
                    {
                        lstProjectItem.Add(projectItem);
                    }
                }
                else
                {
                    logger.Info("Finish convert Job : count=" + lstProjectItem.Count);
                    return;
                }
                iRowIndex++;
                itemId++;
            }
        }
        //更新標單次序
        private TND_PROJECT_ITEM convertRow2TndProjectItem4Refresh(int id, IRow row, int excelrow)
        {
            TND_PROJECT_ITEM projectItem = new TND_PROJECT_ITEM();
            projectItem.PROJECT_ID = projId;
            logger.Info("excelrow=" + excelrow + ",row count=" + row.Cells.Count + ",phyColumn=" + row.PhysicalNumberOfCells);
            string strVal = "";
            int i = 0;
            foreach (var c in row.Cells)
            {
                strVal = strVal + "[" + i + "]=" + c.ToString();
                i++;
            }
            logger.Debug("Excel Row=" + excelrow + ",value=" + strVal);
            //至少需要兩個Column
            if (row.Cells[0].ToString().Trim() != "")//PK 0
            {
                projectItem.PROJECT_ITEM_ID = row.Cells[0].ToString();
                projectItem.EXCEL_ROW_ID = long.Parse(row.Cells[row.Cells.Count - 2].ToString());
                projectItem.CREATE_DATE = System.DateTime.Now;
            }
            else
            {
                return null;
            }
            //假設僅作為排版
            if (row.PhysicalNumberOfCells == 2)
            {
                return projectItem;
            }
            if (row.Cells[1].ToString().Trim() != "")//項次 1
            {
                projectItem.ITEM_ID = row.Cells[1].ToString();
            }
            if (row.Cells[2].ToString().Trim() != "")//名稱 2
            {
                projectItem.ITEM_DESC = row.Cells[2].ToString();
            }

            if (row.Cells[3].ToString().Trim() != "")//單位 3
            {
                projectItem.ITEM_UNIT = row.Cells[3].ToString();
            }
            if (row.Cells[4].ToString().Trim() != "")//數量 4
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[4].ToString());
                    projectItem.ITEM_QUANTITY = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Item_Desc= " + projectItem.ITEM_DESC + ",value=" + row.Cells[4].ToString());
                    logger.Error(e.Message);
                }

            }
            if (row.Cells[5].ToString().Trim() != "")//圖算數量 5
            {
            }
            //欄位不足跳脫處理
            if (row.Cells.Count < 7)
            {
                logErrorMessage("data format Error on ExcelRow=" + excelrow + " column not enough");
                return projectItem;
            }
            if (row.Cells[6].ToString().Trim() != "")//單價  6
            {
                try
                {
                    decimal dPrice = decimal.Parse(row.Cells[6].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[4].ToString());
                    projectItem.ITEM_UNIT_PRICE = dPrice;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Item_Desc= " + projectItem.ITEM_DESC + ",value=" + row.Cells[6].ToString());
                    logger.Error(e.StackTrace);
                }

            }

            if (row.Cells[7].ToString().Trim() != "")//複價 7
            {

            }

            if (row.Cells[8].ToString().Trim() != "")//備註 8
            {
                projectItem.ITEM_REMARK = row.Cells[8].ToString();
            }
            if (row.PhysicalNumberOfCells > 13)
            {
                if (row.Cells[9].ToString().Trim() != "")//九宮格 9
                {
                    projectItem.TYPE_CODE_1 = row.Cells[9].ToString();
                }
                if (row.Cells[10].ToString().Trim() != "")//次九宮格 10
                {
                    projectItem.TYPE_CODE_2 = row.Cells[10].ToString();
                }
                if (row.Cells[11].ToString().Trim() != "")//主系統 11
                {
                    projectItem.SYSTEM_MAIN = row.Cells[11].ToString();
                }
                if (row.Cells[12].ToString().Trim() != "")//次系統 12
                {
                    projectItem.SYSTEM_SUB = row.Cells[12].ToString();
                }
            }
            else
            {
                int cellNo = row.Cells.Count;
                if (row.Cells[cellNo - 5].ToString().Trim() != "")//九宮格 9
                {
                    projectItem.TYPE_CODE_1 = row.Cells[cellNo - 5].ToString();
                }
                if (row.Cells[cellNo - 4].ToString().Trim() != "")//次九宮格 10
                {
                    projectItem.TYPE_CODE_2 = row.Cells[cellNo - 4].ToString();
                }
                if (row.Cells[cellNo - 2 - 3].ToString().Trim() != "")//主系統 11
                {
                    projectItem.SYSTEM_MAIN = row.Cells[cellNo - 3].ToString();
                }
                if (row.Cells[cellNo - 2].ToString().Trim() != "")//次系統 12
                {
                    projectItem.SYSTEM_SUB = row.Cells[cellNo - 2].ToString();
                }
                logger.Debug("excel row:" + excelrow + ",type_code1=" + projectItem.TYPE_CODE_1 + ",type_code2=" + projectItem.TYPE_CODE_2 + ",system_main=" + projectItem.SYSTEM_MAIN + ",system_sub=" + projectItem.SYSTEM_SUB);
            }
            logger.Info("TndprojectItem=" + projectItem.ToString());
            return projectItem;
        }
    }
    #endregion
}

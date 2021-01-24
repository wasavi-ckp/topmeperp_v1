using log4net;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    public class poi4CostChangeService : ExcelBase
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public TND_PROJECT project = null;
        public PLAN_COSTCHANGE_FORM costChangeForm = null;
        public List<PLAN_COSTCHANGE_ITEM> lstItem = null;
        protected SYS_USER user = null;

        public poi4CostChangeService()
        {
            //定義樣板檔案名稱
            templateFile = strUploadPath + "\\CostChange_Template_v1.xlsx";
            logger.Debug("Constroctor!!" + templateFile);
        }
        public void setUser(SYS_USER u)
        {
            user = u;
        }
        public void downloadTemplate()
        {

        }
        /// <summary>
        /// 下載異動單資料
        /// </summary>
        /// <param name="project"></param>                  
        public void createExcel(TND_PROJECT project)
        {
            InitializeWorkbook();
            SetOpSheet("異動單");
            //填寫專案資料
            IRow row = sheet.GetRow(1);
            row.Cells[1].SetCellValue(project.PROJECT_ID);
            row.Cells[2].SetCellValue(project.PROJECT_NAME);

            //填入明細資料
            //ConverObjectTotExcel(lstItem, 4);
            //令存新檔至專案所屬目錄
            outputFile = strUploadPath + "\\" + project.PROJECT_ID + "\\" + project.PROJECT_ID + "-CostChange.xlsx";
            logger.Debug("export excel file=" + outputFile);
            var file = new FileStream(outputFile, FileMode.Create);
            logger.Info("output file=" + file.Name);
            hssfworkbook.Write(file);
            file.Close();
        }
        public void createExcel(TND_PROJECT project, PLAN_COSTCHANGE_FORM form, List<PLAN_COSTCHANGE_ITEM> lstItem)
        {
            InitializeWorkbook();
            SetOpSheet("異動單");
            //填寫專案資料
            IRow row = sheet.GetRow(1);
            row.Cells[1].SetCellValue(project.PROJECT_ID);
            row.Cells[2].SetCellValue(project.PROJECT_NAME);
            //填入異動單資料
            row = sheet.GetRow(2);
            row.Cells[1].SetCellValue(form.FORM_ID);
            row.Cells[3].SetCellValue(form.REMARK_ITEM);
            //填入明細資料
            ConverObjectTotExcel(lstItem, 4);
            //令存新檔至專案所屬目錄
            outputFile = strUploadPath + "\\" + project.PROJECT_ID + "\\" + project.PROJECT_ID + "-" + form.FORM_ID + "_CostChange.xlsx";
            logger.Debug("export excel file=" + outputFile);
            var file = new FileStream(outputFile, FileMode.Create);
            logger.Info("output file=" + file.Name);
            hssfworkbook.Write(file);
            file.Close();
        }
        //轉換物件
        public void ConverObjectTotExcel(List<PLAN_COSTCHANGE_ITEM> lstItem, int startrow)
        {
            int idxRow = 4;

            foreach (PLAN_COSTCHANGE_ITEM item in lstItem)
            {
                logger.Debug("Row Id=" + idxRow + "," + item.ITEM_DESC);
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                //編號 標單編號 項次 品項名稱 單位 單價 異動數量 備註說明 轉入標單
                row.CreateCell(0).SetCellValue(item.ITEM_UID);//PK(PROJECT_ITEM_ID)
                row.Cells[0].CellStyle = style;
                row.CreateCell(1).SetCellValue(item.PLAN_ITEM_ID);//標單編號
                row.Cells[1].CellStyle = style;
                row.CreateCell(2).SetCellValue(item.ITEM_ID);//項次
                row.Cells[2].CellStyle = style;
                row.CreateCell(3).SetCellValue(item.ITEM_DESC);// 品項名稱
                row.Cells[3].CellStyle = style;
                row.CreateCell(4).SetCellValue(item.ITEM_UNIT);//單位
                row.Cells[4].CellStyle = style;

                //異動數量
                if (null != item.ITEM_QUANTITY && item.ITEM_QUANTITY.ToString().Trim() != "")
                {
                    row.CreateCell(5).SetCellValue(double.Parse(item.ITEM_QUANTITY.ToString()));
                    row.Cells[5].CellStyle = style;
                }
                else
                {
                    row.CreateCell(5).SetCellValue("");
                }

                //單價 (還沒決定)
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
                //複價
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
                //8 備註
                if (null != item.ITEM_REMARK && item.ITEM_REMARK.ToString().Trim() != "")
                {
                    row.CreateCell(8).SetCellValue(item.ITEM_REMARK);
                    row.Cells[8].CellStyle = style;
                }
                else
                {
                    row.CreateCell(8).SetCellValue("");
                }
                //9 追加/轉入標單
                if (null != item.TRANSFLAG && item.TRANSFLAG.ToString().Trim() != "" && item.TRANSFLAG.ToString() == "1")
                {
                    row.CreateCell(9).SetCellValue("Y");
                    row.Cells[9].CellStyle = style;
                }
                else
                {
                    row.CreateCell(9).SetCellValue("N");
                }

                idxRow++;
            }
            logger.Info("InitialQuotation finish!!");
        }
        //由Excel 讀取資料
        public void getDataFromExcel(string filpath, string projectId, string formId)
        {
            InitializeWorkbook(filpath);
            SetOpSheet("異動單");
            //讀取專案資料
            IRow row = sheet.GetRow(1);
            project = new TND_PROJECT();
            project.PROJECT_ID = projectId;
            project.PROJECT_NAME = row.Cells[2].ToString();
            logger.Debug("project id=" + project.PROJECT_ID + ",project name=" + project.PROJECT_NAME);
            //取得異動單資料
            row = sheet.GetRow(2);
            costChangeForm = new PLAN_COSTCHANGE_FORM();
            costChangeForm.PROJECT_ID = project.PROJECT_ID;
            costChangeForm.FORM_ID = formId;
            //檢查是否為新異動單
            if (null == costChangeForm.FORM_ID || costChangeForm.FORM_ID == "")
            {
                costChangeForm.CREATE_USER_ID = user.USER_ID;
                costChangeForm.CREATE_DATE = DateTime.Now;
            }
            else
            {
                costChangeForm.MODIFY_USER_ID = user.USER_ID;
                costChangeForm.MODIFY_DATE = DateTime.Now;
            }
            //未送審前狀態不變
            costChangeForm.STATUS = "新建立";
            costChangeForm.REMARK_ITEM = row.Cells[3].ToString();
            logger.Debug("FORM id=" + costChangeForm.FORM_ID);
            ConvertExcel2Object();
        }
        //將Excel Row 轉製成物件
        public void ConvertExcel2Object()
        {
            lstItem = new List<PLAN_COSTCHANGE_ITEM>();
            logger.Debug(sheet.LastRowNum);
            for (int idxRow = 4; idxRow < sheet.LastRowNum + 1; idxRow++)
            {
                PLAN_COSTCHANGE_ITEM item = new PLAN_COSTCHANGE_ITEM();
                if (costChangeForm.FORM_ID != "")
                {
                    item.FORM_ID = costChangeForm.FORM_ID;
                    item.PROJECT_ID = project.PROJECT_ID;
                }
                //編號 標單編號 項次 品項名稱 單位 單價 異動數量 備註說明 轉入標單
                IRow row = sheet.GetRow(idxRow);
                ICell cell = row.GetCell(0);
                if (null != cell)
                {
                    string strUid = cell.ToString();
                    try
                    {
                        item.ITEM_UID = long.Parse(strUid);
                    }
                    catch (Exception ex)
                    {
                        logger.Warn("Excel Row=" + idxRow + " not UID;");
                        logger.Error(ex.Message + ":" + ex.StackTrace);
                    }
                }
                cell = row.GetCell(1);
                if (null != cell)
                {
                    item.PLAN_ITEM_ID = cell.ToString();//標單編號
                }

                cell = row.GetCell(2);
                if (null != cell)
                {
                    item.ITEM_ID = cell.ToString();//標單像次
                }
                cell = row.GetCell(3);
                if (null != cell)
                {
                    item.ITEM_DESC = cell.ToString();// 品項名稱
                }
                cell = row.GetCell(4);
                if (null != cell)
                {
                    item.ITEM_UNIT = cell.ToString();
                }
                logger.Debug("Row Id=" + idxRow + "," + item.ITEM_DESC);
                //異動數量
                cell = row.GetCell(5);
                string strQty = "";
                if (null != cell)
                {
                    strQty = cell.ToString();
                }
                try
                {
                    item.ITEM_QUANTITY = long.Parse(strQty);
                }
                catch (Exception ex)
                {
                    logger.Warn("Excel Row=" + idxRow + " Qty can not covert");
                    logger.Error(ex.Message + ":" + ex.StackTrace);
                }

                //異動單價
                cell = row.GetCell(6);
                string strPrice = "";
                if (null != cell)
                {
                    strPrice = cell.ToString();
                    try
                    {
                        item.ITEM_UNIT_PRICE = long.Parse(strPrice);
                    }
                    catch (Exception ex)
                    {
                        logger.Warn("Excel Row=" + idxRow + " Priec can not covert");
                        logger.Error(ex.Message + ":" + ex.StackTrace);
                    }
                }
                //複價7

                //8 備註
                //異動數量
                cell = row.GetCell(8);
                if (null != cell)
                {
                    item.ITEM_REMARK = cell.ToString();
                }
                //9 追加/轉入標單row.Cells[8].ToString();
                cell = row.GetCell(9);
                string strTransFlag = "Y";
                if (null != cell)
                {
                    strTransFlag = cell.ToString();
                }

                if (null != strTransFlag && strTransFlag != "" && strTransFlag != "N")
                {
                    item.TRANSFLAG = "1";
                }
                else
                {
                    item.TRANSFLAG = "0";
                }
                lstItem.Add(item);
            }
        }
    }
}

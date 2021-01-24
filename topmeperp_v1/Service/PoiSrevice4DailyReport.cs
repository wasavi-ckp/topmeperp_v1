using ICSharpCode.SharpZipLib.Zip;
using log4net;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using topmeperp.Models;


namespace topmeperp.Service
{
    public class SummaryDailyReportToExcel : ExpenseFormToExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string expenseFile = ContextService.strUploadPath + "\\dailyReport_Summary.xlsx";
        string outputPath = ContextService.strUploadPath;
        TND_PROJECT p = null;
        List<SummaryDailyReport> dailyReport = null;
        ISheet sheet = null;

        //建立預算下載表格
        public string exportExcel(List<SummaryDailyReport> _dailyReport)
        {
            //1.讀取費用表格檔案
            InitializeWorkbook(expenseFile);

            if (null != _dailyReport)
            {
                dailyReport = _dailyReport;
                sheet = (XSSFSheet)hssfworkbook.GetSheet("日報彙整表");

                PlanService s = new PlanService();
                p = s.getProject(_dailyReport[0].PROJECT_ID);
                //填寫檔頭
                writeHead();
                //寫入估驗明細資料
                writeEstimateItem();
                //4.另存新檔至專案所屬目錄 (增加Temp for zip 打包使用
                string fileLocation = null;
                fileLocation = outputPath + "\\" + dailyReport[0].PROJECT_ID + "\\" + dailyReport[0].PROJECT_ID + "_" + dailyReport[0].REPORT_START_DATE.Value.ToString("yyyyMM") + "_施工日報.xlsx";
                var file = new FileStream(fileLocation, FileMode.Create);
                logger.Info("new file name =" + file.Name + ",path=" + file.Position);
                hssfworkbook.Write(file);
                file.Close();
                return fileLocation;
            }
            return null;
        }
        //寫入檔頭資料
        protected void writeHead()
        {
            //1.填入表頭資料
            logger.Debug("Table Head_1=" + sheet.GetRow(1).Cells[0].ToString());
            sheet.GetRow(0).Cells[4].SetCellValue("協成水電工程事業有限公司 施工日報彙整");
            logger.Debug("Row 1 count:" + sheet.GetRow(1).Cells.Count());
            for (int i = 0; i < 21; i++)
            {
                logger.Debug("Cell(" + i + ")=" + sheet.GetRow(1).GetCell(i));
            }
            sheet.GetRow(1).Cells[1].SetCellValue(p.PROJECT_NAME + "(" + p.PROJECT_ID + ")");
            sheet.GetRow(1).Cells[10].SetCellValue(dailyReport[0].BGN_REPORT_ID + "~" + dailyReport[0].END_REPORT_ID);
            sheet.GetRow(1).Cells[17].SetCellValue(dailyReport[0].REPORT_START_DATE.Value.ToString("yyyy/MM/dd") + "~" + dailyReport[0].REPORT_END_DATE.Value.ToString("yyyy/MM/dd"));
        }
        //將日報廠商與施作品項寫入Sheet 內
        protected void writeEstimateItem()
        {
            int iStartRow = 4;
            IRow rowSource = sheet.GetRow(iStartRow);
            //iStartRow++;
            string strSupplier = "";
            double amount4supplier = 0;
            foreach (SummaryDailyReport dr in dailyReport)
            {
                IRow rowInsert = sheet.CreateRow(iStartRow);
                rowInsert.Height = rowSource.Height;

                if (strSupplier != "" && strSupplier != dr.SUPPLIER_ID)
                {
                    logger.Debug("Create Summary Row:" + iStartRow);
                    //補上小計資料
                    //rowInsert = sheet.CreateRow(iStartRow);
                    //rowInsert.CreateCell(0).SetCellValue(strSupplier);
                    rowInsert.CreateCell(12).SetCellValue("小計");
                    rowInsert.CreateCell(13).SetCellValue(amount4supplier);
                    //strSupplier = dr.SUPPLIER_ID;
                    amount4supplier = 0.0;
                    iStartRow++;
                    rowInsert = sheet.CreateRow(iStartRow);
                }
                if (!strSupplier.Equals(dr.SUPPLIER_ID))
                {
                    //廠商資料
                    logger.Debug("Create Supplier Row:" + iStartRow);
                    strSupplier = dr.SUPPLIER_ID;
                    rowInsert.CreateCell(0).SetCellValue(strSupplier);
                    amount4supplier = amount4supplier + FillRow(iStartRow, dr, rowInsert);
                }
                else
                {
                    //logger.Debug("Create Item Row:" + iStartRow);
                    logger.Debug("Create Item Row :" + dr.PROJECT_ITEM_ID + " ,iRow=" + iStartRow);
                    rowInsert.CreateCell(0).SetCellValue("");
                    amount4supplier = amount4supplier + FillRow(iStartRow, dr, rowInsert);
                }

                iStartRow++;
            }
            //寫入簽核欄位
            logger.Debug("create row for approve field:" + iStartRow);
            IRow rowEnd = sheet.CreateRow(iStartRow);
            rowEnd.CreateCell(0).SetCellValue("核准");
            rowEnd.CreateCell(4).SetCellValue("複審");
            rowEnd.CreateCell(9).SetCellValue("初審");
            rowEnd.CreateCell(13).SetCellValue("承辦人");
        }

        private static double FillRow(int iStartRow, SummaryDailyReport dr, IRow rowInsert)
        {
            //合約項次
            rowInsert.CreateCell(1).SetCellValue(dr.ITEM_ID);
            //合約品名
            rowInsert.CreateCell(2).SetCellValue(dr.ITEM_DESC);
            //單位
            rowInsert.CreateCell(8).SetCellValue(dr.ITEM_UNIT);
            //前期數量
            rowInsert.CreateCell(10).SetCellValue((double)dr.ACCUMULATE_QTY.Value);
            //前期金額
            rowInsert.CreateCell(11).SetCellFormula("K" + (iStartRow + 1) + "*" + dr.UNIT_COST);
            double amount4supplier = (double)dr.QTY * (double)dr.UNIT_COST;
            //本期數量
            rowInsert.CreateCell(12).SetCellValue((double)dr.QTY);
            //本期金額
            rowInsert.CreateCell(13).SetCellFormula("M" + (iStartRow + 1) + "*" + dr.UNIT_COST);
            //累計數量
            rowInsert.CreateCell(14).SetCellFormula("K" + (iStartRow + 1) + "+M" + (iStartRow + 1));
            //累計金額
            rowInsert.CreateCell(15).SetCellFormula("L" + (iStartRow + 1) + "+N" + (iStartRow + 1));
            //比率
            rowInsert.CreateCell(16).SetCellValue("");
            //備註
            rowInsert.CreateCell(17).SetCellValue(dr.PROJECT_ITEM_ID);
            return amount4supplier;
        }

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
}

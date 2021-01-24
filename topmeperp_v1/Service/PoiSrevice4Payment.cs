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
    public class PaymentExpenseFormToExcel : ExpenseFormToExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string expenseFile = ContextService.strUploadPath + "\\expense_form.xlsx";
        string outputPath = ContextService.strUploadPath;
        ContractModels constract = null;
        //記錄整個檔案的Sheet 
        List<String> lstSheetName = new List<String>();
        ISheet sheet = null;
        //分頁Sheet
        int page = 1;

        //建立預算下載表格
        public string exportExcel(ContractModels _constract)
        {
            //1.讀取費用表格檔案
            InitializeWorkbook(expenseFile);

            if (null != _constract)
            {
                constract = _constract;
                lstSheetName.Add("費用表");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("費用表");
                writeHead();
                //寫入估驗明細資料
                writeEstimateItem();
                //寫入憑證明細資料
                writeInvoiceData();
                //寫入總計資料
                writeSumData();
                //4.另存新檔至專案所屬目錄 (增加Temp for zip 打包使用
                string fileLocation = null;
                fileLocation = outputPath + "\\" + constract.project.PROJECT_ID + "\\" + constract.project.PROJECT_ID + "_" + constract.planEST.EST_FORM_ID + "_費用表.xlsx";
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
            sheet.GetRow(0).Cells[4].SetCellValue("協成水電工程事業有限公司 估驗申請單");
            sheet.GetRow(1).Cells[3].SetCellValue(constract.project.PROJECT_NAME);//專案名稱
            sheet.GetRow(0).Cells[20].SetCellValue("第 " + page + " 頁");
            //todo 需要另外處理
            sheet.GetRow(0).Cells[22].SetCellValue("共" + page + "頁");//總頁數?
            sheet.GetRow(1).Cells[9].SetCellValue(constract.planEST.EST_FORM_ID);//費用單編號
            sheet.GetRow(1).Cells[14].SetCellValue(constract.supplier.COMPANY_NAME);//請款人
            sheet.GetRow(1).Cells[18].SetCellValue(constract.planEST.PAYMENT_DATE.Value.ToString("yyyy/MM"));//費用發生年月
            sheet.GetRow(1).Cells[22].SetCellValue(constract.planEST.CREATE_DATE.Value.ToString("yyyy/MM/dd"));//請款日期
        }
        //將估驗單品項寫入Sheet 內
        protected void writeEstimateItem()
        {
            //2.填入項次資料 每15 筆換頁
            //缺換頁
            int intSheetInx = 1;
            sheet = (XSSFSheet)hssfworkbook.GetSheetAt(intSheetInx);

            int idxRow = 4;
            foreach (EstimationItem item in constract.EstimationItems)
            {
                //將寫入資料換入next sheet
                if (idxRow>20)
                {
                    idxRow = 4;
                    intSheetInx++;
                    sheet = (XSSFSheet)hssfworkbook.GetSheetAt(intSheetInx);
                }
                IRow row = sheet.GetRow(idxRow);//.CreateRow(idxRow);
                logger.Info("Row Id=" + idxRow);
                //項次、品名/摘要、單位、預算(合約)金額、前期累計金額、本期金額、累計金額、累計金額比率、會計名稱
                row.Cells[0].SetCellValue(idxRow - 3);
                //品名或摘要
                logger.Debug("ITEM_DESC=" + item.ITEM_DESC);
                row.Cells[1].SetCellValue(item.ITEM_DESC);
                //單位
                row.Cells[7].SetCellValue(item.ITEM_UNIT);
                //預算(合約)金額
                double contractAmt = double.Parse((item.ITEM_QUANTITY * item.ITEM_UNIT_PRICE).ToString());
                row.Cells[9].SetCellValue(contractAmt);
                //前期累計數量

                //前期累計金額

                //本期數量
                if (null != item.EstimationQty && item.EstimationQty.ToString().Trim() != "")
                {
                    row.Cells[14].SetCellValue(double.Parse(item.EstimationQty.ToString()));
                }
                else
                {
                    row.Cells[14].SetCellValue("");
                }
                //本期金額
                if (null != item.EstimationAmount && item.EstimationAmount.ToString().Trim() != "")
                {
                    row.Cells[16].SetCellValue(double.Parse(item.EstimationAmount.ToString()));
                }
                else
                {
                    row.Cells[16].SetCellValue("");
                }
                //累計數量
                double priorQty = double.Parse(((item.PriorQty == null ? 0 : item.PriorQty) + item.EstimationQty).ToString());
                row.Cells[17].SetCellValue(priorQty);
                //累計金額
                double priorAmount = double.Parse((((item.PriorQty == null ? 0 : item.PriorQty) + item.EstimationQty) * item.ITEM_UNIT_PRICE).ToString());
                row.Cells[18].SetCellValue(priorAmount);

                //累計占比 
                double priorRatio = double.Parse((((item.PriorQty == null ? 0 : item.PriorQty) + item.EstimationQty) / item.ITEM_QUANTITY).ToString());
                row.Cells[19].SetCellValue(priorRatio.ToString("#0.##%"));
                //寫入標單的項次資料 (項次欄位很多空白，同時填入Plan_item_id)
                string remark = item.ITEM_ID?? "";
                logger.Debug("ITEM_ID=" + remark);
                row.Cells[20].SetCellValue(remark  +"(" + item.PLAN_ITEM_ID + ")");
                idxRow++;
            }
        }
        //寫入憑證(發票)資料
        protected void writeInvoiceData()
        {
            //Sheet 僅能存放三張憑證
            int intSheetInx = 1;
            sheet = (XSSFSheet)hssfworkbook.GetSheetAt(intSheetInx);
            int iRow = 21;
            foreach (var item in constract.EstimationInvoices)
            {
                //憑證號碼
                sheet.GetRow(iRow).Cells[19].SetCellValue(item.INVOICE_NUMBER);
                //憑證金額
                sheet.GetRow(iRow).Cells[22].SetCellValue(Convert.ToDouble(item.AMOUNT));
                //代號
                sheet.GetRow(iRow).Cells[23].SetCellValue(item.TYPE);
                iRow++;
                if (iRow == 23)
                {
                    iRow = 21;
                    intSheetInx++;
                    sheet = (XSSFSheet)hssfworkbook.GetSheetAt(intSheetInx);
                }
            }
        }
        //寫入彙總資料
        protected void writeSumData()
        {
            //前期金額
            //本期金額
            var cur_payment = 0.0;
            foreach (var um in constract.EstimationItems)
            {
                cur_payment = cur_payment + Convert.ToDouble(um.EstimationAmount);
            }

            //@**本期金額*@
            //@*代付支出 D *@
            sheet.GetRow(22).Cells[3].SetCellValue(Convert.ToDouble(constract.planEST.PAYMENT_TRANSFER));
            //@*外勞費用 E *@
            sheet.GetRow(22).Cells[4].SetCellValue(Convert.ToDouble(constract.planEST.FOREIGN_PAYMENT));
            //@*小計 F *@
            sheet.GetRow(22).Cells[5].CellFormula = "(D23+D24)";
            //@*保留款 H *@
            sheet.GetRow(22).Cells[7].SetCellValue(Convert.ToDouble(constract.planEST.RETENTION_PAYMENT));
            //@*預付款 J*@
            sheet.GetRow(22).Cells[9].SetCellValue(Convert.ToDouble(constract.planEST.PREPAY_AMOUNT));
            //@*保證金不見了  L(11) *@
            //@*暫借款不見了  M(12) *@
            //@*代付扣回 N (13) *@
            sheet.GetRow(22).Cells[13].SetCellValue(Convert.ToDouble(constract.planEST.PAYMENT_DEDUCTION));
            //@*其他扣款 O (14)*@
            sheet.GetRow(22).Cells[14].SetCellValue(Convert.ToDouble(constract.planEST.OTHER_PAYMENT));
            //@**應付金額 Q (16)*@
            sheet.GetRow(22).Cells[16].CellFormula = "(D23+D24)";
            //var cur_payable = (constract.planEST.OTHER_PAYMENT ?? 0.0) + cur_sub_amount - cur_retention - cur_advance - cur_refund - cur_other;
            //  @**稅金*@
            sheet.GetRow(22).Cells[17].SetCellValue(Convert.ToDouble(constract.planEST.TAX_AMOUNT));
            //實付金額
            // var cur_payamount = cur_payment + Convert.ToDouble(cur_payable + cur_tax);

            //合計
            // sheet.GetRow(30).Cells[22].CellFormula = "(W28+W29+W30)"; 
        }
        //寫入代付支出
        protected void writeEstimationHoldPayments()
        {
            int iRow = 27;
            int intSheetInx = 1;
            sheet = (XSSFSheet)hssfworkbook.GetSheetAt(intSheetInx);
            foreach (PLAN_ESTIMATION_HOLDPAYMENT item in constract.EstimationHoldPayments)
            {
                //廠商名稱							
                sheet.GetRow(iRow).Cells[0].SetCellValue(item.SUPPLIER_ID);
                //統一編號		
                sheet.GetRow(iRow).Cells[1].SetCellValue("");
                //代付金額 
                sheet.GetRow(iRow).Cells[2].SetCellValue(Convert.ToDouble(item.HOLD_AMOUNT));
                //說明事項
                sheet.GetRow(iRow).Cells[3].SetCellValue(item.REMARK);
                iRow++;
                if (iRow == 30)
                {
                    iRow = 27;
                    intSheetInx++;
                    sheet = (XSSFSheet)hssfworkbook.GetSheetAt(intSheetInx);
                }
            }
        }
        //TODO : 其他扣款明細
        protected void writeOtherHoldPayments()
        {

        }
        //寫入代付扣回明細
        protected void writeHold4DeductForm()
        {
            int intSheetInx = 1;
            sheet = (XSSFSheet)hssfworkbook.GetSheetAt(intSheetInx);
            int iRow = 34;
            foreach (Model4PaymentTransfer item in constract.Hold4DeductForm)
            {
                //憑單編號 
                sheet.GetRow(iRow).Cells[0].SetCellValue(item.EST_FORM_ID);
                //次數 
                sheet.GetRow(iRow).Cells[1].SetCellValue(Convert.ToDouble(item.ESTIMATION_COUNT));
                //估驗日期 
                sheet.GetRow(iRow).Cells[1].SetCellValue(item.CREATE_DATE);
                //請款人 
                sheet.GetRow(iRow).Cells[1].SetCellValue(item.COMPANY_NAME);
                //代付金額 
                sheet.GetRow(iRow).Cells[1].SetCellValue(Convert.ToDouble(item.PAID_AMOUNT));
                //手續費 
                sheet.GetRow(iRow).Cells[1].SetCellValue(Convert.ToDouble(item.FEE));
                //應扣金額 
                sheet.GetRow(iRow).Cells[1].SetCellValue(Convert.ToDouble(item.HOLD_AMOUNT));
                //本期扣款 
                sheet.GetRow(iRow).Cells[1].SetCellValue(Convert.ToDouble(item.CUR_HOLDAMOUNT));
                //應扣未扣 
                //sheet.GetRow(iRow).Cells[1].SetCellValue(item.EST_FORM_ID);
                //扣款原因
                sheet.GetRow(iRow).Cells[1].SetCellValue(item.REMARK);
                iRow++;
                //寫入5筆資料換頁
                if (iRow == 39)
                {
                    iRow = 34;
                    intSheetInx++;
                    sheet = (XSSFSheet)hssfworkbook.GetSheetAt(intSheetInx);
                }
            }
        }

        //建立新的Sheet for 寫入
        protected ISheet createSheet(int idx)
        {
            ISheet targetSheet = (ISheet)hssfworkbook.CloneSheet(0);
            string strSheetName = "費用表_" + idx;
            if (hssfworkbook.GetSheet(strSheetName)!=null )
            {
                targetSheet= hssfworkbook.GetSheet(strSheetName);
            }
            else
            {
                logger.Info("creare Sheet");
                ISheet tempSheet = hssfworkbook.GetSheet("樣本");

                string strNewSheetName = "費用表_" + lstSheetName.Count;
                targetSheet= tempSheet.CopySheet(strNewSheetName);
            }
            return targetSheet;
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


import java.text.ParseException;
import java.util.List;

public class TestSheetOutput_ConditionalFormatSpaningColumns extends TalendFakeJob {

	/**
	 * @param args
	 * @throws ParseException
	 */
	public static void main(String[] args) throws Exception {
		test_conditionalFormatsSpanningColumns();
	}


	public static class row1Struct {
		final static byte[] commonByteArrayLock_COMPDEV_test21_wide_conditional_formats = new byte[0];
		static byte[] commonByteArray_COMPDEV_test21_wide_conditional_formats = new byte[0];

		public String Employer;

		public String getEmployer() {
			return this.Employer;
		}

		public String Company_Launch_Date;

		public String getCompany_Launch_Date() {
			return this.Company_Launch_Date;
		}

		public String Zebit_ID;

		public String getZebit_ID() {
			return this.Zebit_ID;
		}

		public String Customer_Name;

		public String getCustomer_Name() {
			return this.Customer_Name;
		}

		public java.util.Date Date_Registered;

		public java.util.Date getDate_Registered() {
			return this.Date_Registered;
		}

		public String Current_Account_Status;

		public String getCurrent_Account_Status() {
			return this.Current_Account_Status;
		}

		public String Current_Active_Payment_Method;

		public String getCurrent_Active_Payment_Method() {
			return this.Current_Active_Payment_Method;
		}

		public String RISA_State;

		public String getRISA_State() {
			return this.RISA_State;
		}

		public java.util.Date Date_of_1st_Order;

		public java.util.Date getDate_of_1st_Order() {
			return this.Date_of_1st_Order;
		}

		public java.util.Date Date_of_Last_Order_Placed;

		public java.util.Date getDate_of_Last_Order_Placed() {
			return this.Date_of_Last_Order_Placed;
		}

		public Integer Total_Orders_Placed;

		public Integer getTotal_Orders_Placed() {
			return this.Total_Orders_Placed;
		}

		public Integer Total_ZL_Orders;

		public Integer getTotal_ZL_Orders() {
			return this.Total_ZL_Orders;
		}

		public Integer Total_PIF_Orders;

		public Integer getTotal_PIF_Orders() {
			return this.Total_PIF_Orders;
		}

		public Integer Total_Orders_Cancelled;

		public Integer getTotal_Orders_Cancelled() {
			return this.Total_Orders_Cancelled;
		}

		public Integer Annual_Pay_Frequency;

		public Integer getAnnual_Pay_Frequency() {
			return this.Annual_Pay_Frequency;
		}

		public String ZL_Term;

		public String getZL_Term() {
			return this.ZL_Term;
		}

		public String ZL_Payments_Stream;

		public String getZL_Payments_Stream() {
			return this.ZL_Payments_Stream;
		}

		public String ZL_Amount;

		public String getZL_Amount() {
			return this.ZL_Amount;
		}

		public String ZL_Used;

		public String getZL_Used() {
			return this.ZL_Used;
		}

		public String ZL_Left;

		public String getZL_Left() {
			return this.ZL_Left;
		}

		public String ZL_Initial_Value;

		public String getZL_Initial_Value() {
			return this.ZL_Initial_Value;
		}

		public String ZL_Value_Left;

		public String getZL_Value_Left() {
			return this.ZL_Value_Left;
		}

		public String ZL_Value_Used;

		public String getZL_Value_Used() {
			return this.ZL_Value_Used;
		}

		public String __ZL_Utilization;

		public String get__ZL_Utilization() {
			return this.__ZL_Utilization;
		}

		public String Total_ZL_Amount_Collected__includes_refunds_;

		public String getTotal_ZL_Amount_Collected__includes_refunds_() {
			return this.Total_ZL_Amount_Collected__includes_refunds_;
		}

		public String Total_ZL_Amount_Refunded;

		public String getTotal_ZL_Amount_Refunded() {
			return this.Total_ZL_Amount_Refunded;
		}

		public String Total_ZL_Outstanding_to_Pay;

		public String getTotal_ZL_Outstanding_to_Pay() {
			return this.Total_ZL_Outstanding_to_Pay;
		}

		public String Total_PIF_Collected__includes_refunds_;

		public String getTotal_PIF_Collected__includes_refunds_() {
			return this.Total_PIF_Collected__includes_refunds_;
		}

		public String Total_PIF_Amount_Refunded;

		public String getTotal_PIF_Amount_Refunded() {
			return this.Total_PIF_Amount_Refunded;
		}

		public java.util.Date Next_Payment_Date;

		public java.util.Date getNext_Payment_Date() {
			return this.Next_Payment_Date;
		}

		public String Next_Payment_Amount;

	}

	public static void test_conditionalFormatsSpanningColumns() throws Exception {

		/**
		 * [tFileExcelWorkbookOpen_1 begin ] start
		 */

		ok_Hash.put("tFileExcelWorkbookOpen_1", false);
		start_Hash.put("tFileExcelWorkbookOpen_1", System.currentTimeMillis());

		currentComponent = "tFileExcelWorkbookOpen_1";

		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookOpen_1 = new de.jlo.talendcomp.excel.SpreadsheetFile();
		tFileExcelWorkbookOpen_1.setCreateStreamingXMLWorkbook(false);
		try {
			// read a excel file as template
			// this file file will not used as output file
			tFileExcelWorkbookOpen_1
					.setInputFile("/Volumes/Data/Talend/testdata/excel/test21/Customer_View_Template.xlsx", true);
			tFileExcelWorkbookOpen_1.initializeWorkbook();
		} catch (Exception e) {
			globalMap.put("tFileExcelWorkbookOpen_1_ERROR_MESSAGE", e.getMessage());
			throw e;
		}

		globalMap.put("workbook_tFileExcelWorkbookOpen_1", tFileExcelWorkbookOpen_1.getWorkbook());
		globalMap.put("tFileExcelWorkbookOpen_1_COUNT_SHEETS",
				tFileExcelWorkbookOpen_1.getWorkbook().getNumberOfSheets());

		/**
		 * [tFileExcelWorkbookOpen_1 begin ] stop
		 */

		row1Struct row1 = new row1Struct();

		/**
		 * [tFileExcelSheetOutput_1 begin ] start
		 */

		ok_Hash.put("tFileExcelSheetOutput_1", false);
		start_Hash.put("tFileExcelSheetOutput_1",
				System.currentTimeMillis());

		currentComponent = "tFileExcelSheetOutput_1";

		if (execStat) {
			if (resourceMap.get("inIterateVComp") == null) {

				runStat.updateStatOnConnection("row1" + iterateId, 0, 0);

			}
		}

		de.jlo.talendcomp.excel.SpreadsheetOutput tFileExcelSheetOutput_1 = new de.jlo.talendcomp.excel.SpreadsheetOutput();
		tFileExcelSheetOutput_1.setDebug(true);
		tFileExcelSheetOutput_1
				.setWorkbook((org.apache.poi.ss.usermodel.Workbook) globalMap
						.get("workbook_tFileExcelWorkbookOpen_1"));
		tFileExcelSheetOutput_1.setTargetSheetName("Customer View");
		globalMap.put("tFileExcelSheetOutput_1_SHEET_NAME",
				tFileExcelSheetOutput_1.getTargetSheetName());
		tFileExcelSheetOutput_1.initializeSheet();
		tFileExcelSheetOutput_1.setRowStartIndex(2 - 1);
		tFileExcelSheetOutput_1.setFirstRowIsHeader(false);
		// configure cell positions
		tFileExcelSheetOutput_1.setColumnStart("A");
		tFileExcelSheetOutput_1.setReuseExistingStylesFromFirstWrittenRow(true);
		tFileExcelSheetOutput_1.setReuseFirstRowHeight(false);
		tFileExcelSheetOutput_1
				.setReuseExistingStylesAlternating(false);
		// configure cell formats
		tFileExcelSheetOutput_1.setWriteNullValues(false);
		// configure auto size columns and group rows by and comments
		// row counter
		int nb_line_tFileExcelSheetOutput_1 = 0;

		/**
		 * [tFileExcelSheetOutput_1 begin ] stop
		 */

		/**
		 * [tFixedFlowInput_1 begin ] start
		 */

		ok_Hash.put("tFixedFlowInput_1", false);
		start_Hash.put("tFixedFlowInput_1", System.currentTimeMillis());

		currentComponent = "tFixedFlowInput_1";

		int tos_count_tFixedFlowInput_1 = 0;

		int nb_line_tFixedFlowInput_1 = 0;
		List<row1Struct> cacheList_tFixedFlowInput_1 = new java.util.ArrayList<row1Struct>();
		row1 = new row1Struct();
		row1.Employer = "Jan";
		row1.Company_Launch_Date = null;
		row1.Zebit_ID = null;
		row1.Customer_Name = null;
		row1.Date_Registered = null;
		row1.Current_Account_Status = "Default";
		row1.Current_Active_Payment_Method = null;
		row1.RISA_State = null;
		row1.Date_of_1st_Order = null;
		row1.Date_of_Last_Order_Placed = null;
		row1.Total_Orders_Placed = null;
		row1.Total_ZL_Orders = null;
		row1.Total_PIF_Orders = null;
		row1.Total_Orders_Cancelled = null;
		row1.Annual_Pay_Frequency = null;
		row1.ZL_Term = null;
		row1.ZL_Payments_Stream = null;
		row1.ZL_Amount = null;
		row1.ZL_Used = null;
		row1.ZL_Left = null;
		row1.ZL_Initial_Value = null;
		row1.ZL_Value_Left = null;
		row1.ZL_Value_Used = null;
		row1.__ZL_Utilization = null;
		row1.Total_ZL_Amount_Collected__includes_refunds_ = null;
		row1.Total_ZL_Amount_Refunded = null;
		row1.Total_ZL_Outstanding_to_Pay = null;
		row1.Total_PIF_Collected__includes_refunds_ = null;
		row1.Total_PIF_Amount_Refunded = null;
		row1.Next_Payment_Date = null;
		row1.Next_Payment_Amount = null;
		cacheList_tFixedFlowInput_1.add(row1);
		row1 = new row1Struct();
		row1.Employer = "Hans";
		row1.Company_Launch_Date = null;
		row1.Zebit_ID = null;
		row1.Customer_Name = null;
		row1.Date_Registered = null;
		row1.Current_Account_Status = "Default";
		row1.Current_Active_Payment_Method = null;
		row1.RISA_State = null;
		row1.Date_of_1st_Order = null;
		row1.Date_of_Last_Order_Placed = null;
		row1.Total_Orders_Placed = null;
		row1.Total_ZL_Orders = null;
		row1.Total_PIF_Orders = null;
		row1.Total_Orders_Cancelled = null;
		row1.Annual_Pay_Frequency = null;
		row1.ZL_Term = null;
		row1.ZL_Payments_Stream = null;
		row1.ZL_Amount = null;
		row1.ZL_Used = null;
		row1.ZL_Left = null;
		row1.ZL_Initial_Value = null;
		row1.ZL_Value_Left = null;
		row1.ZL_Value_Used = null;
		row1.__ZL_Utilization = null;
		row1.Total_ZL_Amount_Collected__includes_refunds_ = null;
		row1.Total_ZL_Amount_Refunded = null;
		row1.Total_ZL_Outstanding_to_Pay = null;
		row1.Total_PIF_Collected__includes_refunds_ = null;
		row1.Total_PIF_Amount_Refunded = null;
		row1.Next_Payment_Date = null;
		row1.Next_Payment_Amount = null;
		cacheList_tFixedFlowInput_1.add(row1);
		row1 = new row1Struct();
		row1.Employer = "Eberhard";
		row1.Company_Launch_Date = null;
		row1.Zebit_ID = null;
		row1.Customer_Name = null;
		row1.Date_Registered = null;
		row1.Current_Account_Status = "Frozen";
		row1.Current_Active_Payment_Method = null;
		row1.RISA_State = null;
		row1.Date_of_1st_Order = null;
		row1.Date_of_Last_Order_Placed = null;
		row1.Total_Orders_Placed = null;
		row1.Total_ZL_Orders = null;
		row1.Total_PIF_Orders = null;
		row1.Total_Orders_Cancelled = null;
		row1.Annual_Pay_Frequency = null;
		row1.ZL_Term = null;
		row1.ZL_Payments_Stream = null;
		row1.ZL_Amount = null;
		row1.ZL_Used = null;
		row1.ZL_Left = null;
		row1.ZL_Initial_Value = null;
		row1.ZL_Value_Left = null;
		row1.ZL_Value_Used = null;
		row1.__ZL_Utilization = null;
		row1.Total_ZL_Amount_Collected__includes_refunds_ = null;
		row1.Total_ZL_Amount_Refunded = null;
		row1.Total_ZL_Outstanding_to_Pay = null;
		row1.Total_PIF_Collected__includes_refunds_ = null;
		row1.Total_PIF_Amount_Refunded = null;
		row1.Next_Payment_Date = null;
		row1.Next_Payment_Amount = null;
		cacheList_tFixedFlowInput_1.add(row1);
		for (int i_tFixedFlowInput_1 = 0; i_tFixedFlowInput_1 < 1; i_tFixedFlowInput_1++) {
			for (row1Struct tmpRow_tFixedFlowInput_1 : cacheList_tFixedFlowInput_1) {
				nb_line_tFixedFlowInput_1++;
				row1 = tmpRow_tFixedFlowInput_1;

				/**
				 * [tFixedFlowInput_1 begin ] stop
				 */
				/**
				 * [tFixedFlowInput_1 main ] start
				 */

				currentComponent = "tFixedFlowInput_1";

				tos_count_tFixedFlowInput_1++;

				/**
				 * [tFixedFlowInput_1 main ] stop
				 */

				/**
				 * [tFileExcelSheetOutput_1 main ] start
				 */

				currentComponent = "tFileExcelSheetOutput_1";

				// row1
				// row1

				if (execStat) {
					runStat.updateStatOnConnection("row1" + iterateId,
							1, 1);
				}

				// fill schema data into the object array
				Object[] dataset_tFileExcelSheetOutput_1 = new Object[31];
				dataset_tFileExcelSheetOutput_1[0] = row1.Employer;
				dataset_tFileExcelSheetOutput_1[1] = row1.Company_Launch_Date;
				dataset_tFileExcelSheetOutput_1[2] = row1.Zebit_ID;
				dataset_tFileExcelSheetOutput_1[3] = row1.Customer_Name;
				dataset_tFileExcelSheetOutput_1[4] = row1.Date_Registered;
				dataset_tFileExcelSheetOutput_1[5] = row1.Current_Account_Status;
				dataset_tFileExcelSheetOutput_1[6] = row1.Current_Active_Payment_Method;
				dataset_tFileExcelSheetOutput_1[7] = row1.RISA_State;
				dataset_tFileExcelSheetOutput_1[8] = row1.Date_of_1st_Order;
				dataset_tFileExcelSheetOutput_1[9] = row1.Date_of_Last_Order_Placed;
				dataset_tFileExcelSheetOutput_1[10] = row1.Total_Orders_Placed;
				dataset_tFileExcelSheetOutput_1[11] = row1.Total_ZL_Orders;
				dataset_tFileExcelSheetOutput_1[12] = row1.Total_PIF_Orders;
				dataset_tFileExcelSheetOutput_1[13] = row1.Total_Orders_Cancelled;
				dataset_tFileExcelSheetOutput_1[14] = row1.Annual_Pay_Frequency;
				dataset_tFileExcelSheetOutput_1[15] = row1.ZL_Term;
				dataset_tFileExcelSheetOutput_1[16] = row1.ZL_Payments_Stream;
				dataset_tFileExcelSheetOutput_1[17] = row1.ZL_Amount;
				dataset_tFileExcelSheetOutput_1[18] = row1.ZL_Used;
				dataset_tFileExcelSheetOutput_1[19] = row1.ZL_Left;
				dataset_tFileExcelSheetOutput_1[20] = row1.ZL_Initial_Value;
				dataset_tFileExcelSheetOutput_1[21] = row1.ZL_Value_Left;
				dataset_tFileExcelSheetOutput_1[22] = row1.ZL_Value_Used;
				dataset_tFileExcelSheetOutput_1[23] = row1.__ZL_Utilization;
				dataset_tFileExcelSheetOutput_1[24] = row1.Total_ZL_Amount_Collected__includes_refunds_;
				dataset_tFileExcelSheetOutput_1[25] = row1.Total_ZL_Amount_Refunded;
				dataset_tFileExcelSheetOutput_1[26] = row1.Total_ZL_Outstanding_to_Pay;
				dataset_tFileExcelSheetOutput_1[27] = row1.Total_PIF_Collected__includes_refunds_;
				dataset_tFileExcelSheetOutput_1[28] = row1.Total_PIF_Amount_Refunded;
				dataset_tFileExcelSheetOutput_1[29] = row1.Next_Payment_Date;
				dataset_tFileExcelSheetOutput_1[30] = row1.Next_Payment_Amount;
				// write dataset
				try {
					tFileExcelSheetOutput_1
							.writeRow(dataset_tFileExcelSheetOutput_1);
					nb_line_tFileExcelSheetOutput_1++;
				} catch (Exception e) {
					globalMap.put(
							"tFileExcelSheetOutput_1_ERROR_MESSAGE",
							"Error in line "
									+ nb_line_tFileExcelSheetOutput_1
									+ ":" + e.getMessage());
					throw e;
				}

				/**
				 * [tFileExcelSheetOutput_1 main ] stop
				 */

				/**
				 * [tFixedFlowInput_1 end ] start
				 */

				currentComponent = "tFixedFlowInput_1";

			}
		}
		cacheList_tFixedFlowInput_1.clear();
		globalMap.put("tFixedFlowInput_1_NB_LINE",
				nb_line_tFixedFlowInput_1);

		ok_Hash.put("tFixedFlowInput_1", true);
		end_Hash.put("tFixedFlowInput_1", System.currentTimeMillis());

		/**
		 * [tFixedFlowInput_1 end ] stop
		 */

		/**
		 * [tFileExcelSheetOutput_1 end ] start
		 */

		currentComponent = "tFileExcelSheetOutput_1";

		tFileExcelSheetOutput_1.setupColumnSize();
		tFileExcelSheetOutput_1.closeLastGroup();
		tFileExcelSheetOutput_1
				.extendCellRangesForConditionalFormattings();
		globalMap.put("tFileExcelSheetOutput_1_NB_LINE",
				nb_line_tFileExcelSheetOutput_1);
		globalMap.put("tFileExcelSheetOutput_1_LAST_ROW_INDEX",
				tFileExcelSheetOutput_1.getSheetLastRowIndex() + 1);
		if (execStat) {
			if (resourceMap.get("inIterateVComp") == null
					|| !((Boolean) resourceMap.get("inIterateVComp"))) {
				runStat.updateStatOnConnection("row1" + iterateId, 2, 0);
			}
		}

		ok_Hash.put("tFileExcelSheetOutput_1", true);
		end_Hash.put("tFileExcelSheetOutput_1",
				System.currentTimeMillis());

		/**
		 * [tFileExcelSheetOutput_1 end ] stop
		 */


	}

}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.IO;
using System.Data.Odbc;
using System.Collections;

namespace SQLFinance
{
    public partial class formMain : Form
    {
        public Excel.Application excelApp = new Excel.Application();
        public String connectionString;
        public SqlConnection connection = new SqlConnection("");
        public SqlDataAdapter dataAdapter;
        public OdbcConnection odbcConnection;
        public DataSet cat1DataSet;
        public DataSet newCat1DataSet;
        public DataSet cat2DataSet;
        public DataSet cat3DataSet;
        public DataSet cat4DataSet;
        public DataSet catAllDataSet;
        public DataSet packagingDataSet;
        public DataSet manualInvoiceDataSet;
        public string packagingUsedQuery;
        public string rawsUsedMcQuery;
        public string monthlyInvoicesQuery;
        public string newCat1Query;
        public string onlyMcCormick = "";
        public string cat1cat2Switch = "";
        public string cat1SwitchQuery = "";
        public string cat2SwitchQuery = "";
        public string cat4FullDetailQuery = "";
        public string manualInvoices = "";
        public string manualInvoicesQuery = "";
        public string UN = "";
        public string PW = "";
        public string cat4Numbers = "";
        public bool reallyClose = false;
        public ArrayList cat4InvoiceNumbers = new ArrayList();
        public Excel.Workbook excelWorkbook;

        /* 
         * Sets Max dates to today to avoid whatever nonsense
         * a query into the future would entail
         * 
         * Sets the default start date to the first of 
         * the current month
         * 
         * Establishes connection to SQL database
         * 
         */

        public formMain()
        {
            InitializeComponent();

            dtpStartDate.MaxDate = DateTime.Today;
            dtpStartDate.Value = DateTime.Parse(DateTime.Today.Month + "/1/" + DateTime.Today.Year);
            dtpEndDate.MaxDate = DateTime.Today;
            dtpEndDate.Value = DateTime.Today;

            while (connection.State != ConnectionState.Open && !reallyClose)
            {
                //MessageBox.Show(reallyClose.ToString());
                formLogin loginForm = new formLogin(this);
                loginForm.ShowDialog();

                connectionString = "Data Source=192.168.100.6; Persist Security Info=True;User ID = " + UN + " ; Password = " + PW + ";";
                connection = new SqlConnection(connectionString);

                try
                {
                    connection.Open();
                    MessageBox.Show("Connection Successful!");
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "Connection Unsuccessful");
                    System.Environment.Exit(1);
                }
            }
        }

        private void cmdGo_Click(object sender, EventArgs e)
        {
            /*
             *Function runs whenever the Go button is clicked
             */
            if (dtpStartDate.Value <= dtpEndDate.Value)
            {
                if (txtFileOutput.Text != "")
                {
                    getCat1Data();
                    getCat2Data();
                    getCat3Data();
                    getCatAllData();
                    getCat4Data();
                    getManualInvoiceData();

                    excelWorkbook = excelApp.Workbooks.Add();
                    printToWorksheet("Cat4", cat4DataSet);
                    printManualInvoices();

                    excelApp.Worksheets.Add();
                    printToWorksheet("Cat3", cat3DataSet);

                    excelApp.Worksheets.Add();
                    printToWorksheet("Cat2", cat2DataSet);

                    excelApp.Worksheets.Add();
                    printToWorksheet("Cat1", cat1DataSet);

                    excelApp.Worksheets.Add();
                    printToWorksheet("CatAll", catAllDataSet);

                    formatWorkbook();
                    MessageBox.Show("Save Completed!");
                }
                else
                {
                    MessageBox.Show("Please select a file output location!");
                }
            }
            else 
            {
                MessageBox.Show("Start date must be before end date!");
            }
        }

        private void formatWorkbook() {
            /*
             *Creates a table on the CatAll sheet with a totals row
             */
            Excel.Worksheet ws = excelApp.ActiveSheet;
            Excel.Range tableRange = ws.Range["A1", "E" + (catAllDataSet.Tables[0].Rows.Count + 1)];
            ws.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, tableRange, Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "CatAllTable";
            var catAllTable = ws.ListObjects["CatAllTable"];
            catAllTable.TableStyle = "TableStyleMedium2";
            catAllTable.ShowTotals = true;
            /* 
             *Additional formatting on table columns
             */
            ws.Columns["A:A"].ColumnWidth = 9;
            ws.Columns["B:B"].ColumnWidth = 42.43;
            ws.Columns["C:C"].ColumnWidth = 22.29;
            ws.Columns["D:D"].ColumnWidth = 12.57;
            ws.Columns["D:D"].NumberFormat = "$ #,##0.00";
            catAllTable.ListColumns[4].TotalsCalculation = Excel.XlTotalsCalculation.xlTotalsCalculationSum;
            catAllTable.ListColumns[5].TotalsCalculation = Excel.XlTotalsCalculation.xlTotalsCalculationNone;
            /*
             *Saves the file and closes the workbook
             *Also closes the application which is important
             *for preventing phantom Excel processes in Task Manager
             */
            excelApp.ActiveWorkbook.SaveAs(txtFileOutput.Text);
            excelApp.ActiveWorkbook.Close();
            excelApp.Quit();
        }

        private void printToWorksheet(string name, DataSet dataSet) {
            excelApp.Worksheets[1].Name = name;
            excelApp.Worksheets[name].Select();
            printData(dataSet);
        }

        private void getCat1Data() 
        {
         /*
          *Changes SQL Query string and then dumps the
          *query results in a dataset.
          *Shows dataset in serperate form
          */
            onlyMcCormick = "";
            cat1cat2Switch = cat1SwitchQuery;
            setStrings();
            dataAdapter = new SqlDataAdapter(rawsUsedMcQuery, connection);
            cat1DataSet = new DataSet();
            dataAdapter.Fill(cat1DataSet);

            dataDisplay dataDisplayForm = new dataDisplay(cat1DataSet);
            dataDisplayForm.Show();
        }
        private void getCat2Data()
        {
         /*
          *Changes SQL Query string and then dumps the
          *query results in a dataset.
          *Shows dataset in serperate form
          */
            cat1cat2Switch = cat2SwitchQuery;
            setStrings();
            dataAdapter = new SqlDataAdapter(rawsUsedMcQuery, connection);
            cat2DataSet = new DataSet();
            dataAdapter.Fill(cat2DataSet);
            dataDisplay dataDisplayForm = new dataDisplay(cat2DataSet);
            dataDisplayForm.Show();
        }
        private void getCat3Data()
        {
         /*
          *Changes SQL Query string and then dumps the
          *query results in a dataset.
          *Shows dataset in serperate form
          */
            setStrings();
            dataAdapter = new SqlDataAdapter(packagingUsedQuery, connection);
            cat3DataSet = new DataSet();
            dataAdapter.Fill(cat3DataSet);
            dataDisplay dataDisplayForm = new dataDisplay(cat3DataSet);
            dataDisplayForm.Show();
        }

        private void getCat4Data() 
        {
            setStrings();
            dataAdapter = new SqlDataAdapter(cat4FullDetailQuery, connection);
            cat4DataSet = new DataSet();
            dataAdapter.Fill(cat4DataSet);
            dataDisplay dataDisplayForm = new dataDisplay(cat4DataSet);
            dataDisplayForm.Show();
        }

        private void getManualInvoiceData()
        {
            setStrings();
            dataAdapter = new SqlDataAdapter(manualInvoicesQuery, connection);
            manualInvoiceDataSet = new DataSet();
            dataAdapter.Fill(manualInvoiceDataSet);
            dataDisplay dataDisplayForm = new dataDisplay(manualInvoiceDataSet);
            dataDisplayForm.Show();
        }

        private void getCatAllData()
        { 
         /*
          *Changes SQL Query string and then dumps the
          *query results in a dataset.
          *Shows dataset in serperate form
          */
            setStrings();
            dataAdapter = new SqlDataAdapter(monthlyInvoicesQuery, connection);
            catAllDataSet = new DataSet();
            dataAdapter.Fill(catAllDataSet);
            catAllDataSet.Tables[0].Columns.Add("Category", typeof(string));
            /*
             *Checks if data from previous datasets exists in catAll dataset
             */
            for (int i = 0; i < catAllDataSet.Tables[0].Rows.Count; i++) {
                catAllDataSet.Tables[0].Rows[i][4] = "Cat4";
                for (int j = 0; j < cat1DataSet.Tables[0].Rows.Count; j++) {
                    if (catAllDataSet.Tables[0].Rows[i][2].Equals(cat1DataSet.Tables[0].Rows[j][7]))
                    {
                        catAllDataSet.Tables[0].Rows[i][4] = "Cat1";
                    }
                }
                for (int k = 0; k < cat2DataSet.Tables[0].Rows.Count; k++)
                {
                    if (catAllDataSet.Tables[0].Rows[i][2].Equals(cat2DataSet.Tables[0].Rows[k][7]))
                    {
                        catAllDataSet.Tables[0].Rows[i][4] = "Cat2";
                    }
                }
                for (int l = 0; l < cat3DataSet.Tables[0].Rows.Count; l++)
                {
                    if (catAllDataSet.Tables[0].Rows[i][2].Equals(cat3DataSet.Tables[0].Rows[l][7]))
                    {
                        catAllDataSet.Tables[0].Rows[i][4] = "Cat3";
                    }
                }
            }

            for (int i = 0; i < catAllDataSet.Tables[0].Rows.Count; i++) 
            {
                if (catAllDataSet.Tables[0].Rows[i][4].Equals("Cat4")) 
                {
                    cat4InvoiceNumbers.Add(catAllDataSet.Tables[0].Rows[i][2].ToString());
                    cat4Numbers += "'" + catAllDataSet.Tables[0].Rows[i][2].ToString() + "',";
                }
            }

            cat4Numbers = cat4Numbers.Substring(0, cat4Numbers.Length - 1);
            setStrings();

            dataDisplay dataDisplayForm = new dataDisplay(catAllDataSet);
            dataDisplayForm.Show();
        }

        public void rectifyNewCat1() 
        {
            /*
             *Iterates through newCat1DataSet and checks if any of those lines
             *appear in any other dataset
             *While technically legacy code, it is being left here for
             *future use
             */
            for (int i = 0; i < newCat1DataSet.Tables[0].Rows.Count; i++) 
            {
                for (int j = 0; j < cat2DataSet.Tables[0].Rows.Count; j++) 
                {
                    if (newCat1DataSet.Tables[0].Rows[i][7].Equals(cat2DataSet.Tables[0].Rows[j][7])) 
                    {
                        cat2DataSet.Tables[0].Rows[j].Delete();
                        cat2DataSet.AcceptChanges();
                    }
                }
                for (int k = 0; k < cat3DataSet.Tables[0].Rows.Count; k++)
                {
                    if (newCat1DataSet.Tables[0].Rows[i][7].Equals(cat3DataSet.Tables[0].Rows[k][7]))
                    {
                        cat3DataSet.Tables[0].Rows[k].Delete();
                        cat3DataSet.AcceptChanges();
                    }
                }
            }
        }

        private void printData(DataSet dataSet) {
            /*
             *Iterates through dataset and dumps data onto
             *the Excel worksheet
             */
            for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
            {
                for (int j = 0; j < dataSet.Tables[0].Columns.Count; j++)
                {
                    if (i == 0)
                    {
                        excelWorkbook.ActiveSheet.Cells[i + 1, j + 1] = dataSet.Tables[0].Columns[j].ColumnName;
                        excelWorkbook.ActiveSheet.Cells[i + 2, j + 1] = dataSet.Tables[0].Rows[i][j];
                    }
                    else 
                    {
                        excelWorkbook.ActiveSheet.Cells[i + 2, j + 1] = dataSet.Tables[0].Rows[i][j];
                    }
                    
                }
            }
        }

        private void appendPrintedData(DataSet dataSet1, DataSet dataSet2) 
        {
         /*
          *Checks the amount of rows in dataSet2 and uses
          *that offset to append data from dataSet1 to the end of it.
          *While this is technically legacy code, it is being left here 
          *for future use
          */
            int offset = dataSet2.Tables[0].Rows.Count;
            int k = 0;
            for (int i = offset; i < offset + dataSet1.Tables[0].Rows.Count; i++)
            {
                for (int j = 0; j < dataSet1.Tables[0].Columns.Count; j++)
                {
                    excelWorkbook.ActiveSheet.Cells[i + 2, j + 1] = dataSet1.Tables[0].Rows[k][j];
                }
                k++;
            }
        }

        private void printManualInvoices() 
        {
            int offset = cat4DataSet.Tables[0].Rows.Count;
            int k = 0;
            for (int i = offset; i < offset + manualInvoiceDataSet.Tables[0].Rows.Count; i++)
            {
                excelWorkbook.ActiveSheet.Cells[i + 2, 5] = manualInvoiceDataSet.Tables[0].Rows[k][1];
                excelWorkbook.ActiveSheet.Cells[i + 2, 8] = manualInvoiceDataSet.Tables[0].Rows[k][2];
                excelWorkbook.ActiveSheet.Cells[i + 2, 9] = manualInvoiceDataSet.Tables[0].Rows[k][3];
                excelWorkbook.ActiveSheet.Cells[i + 2, 10] = manualInvoiceDataSet.Tables[0].Rows[k][0];
                k++;
            }
        }

        private void cmdOFD_Click(object sender, EventArgs e)
        {
            /*
             *Uses built in MS Save File Handling to allow user to pick
             *a save location
             */
            SaveFileDialog chooseOutputLocation = new SaveFileDialog();
            chooseOutputLocation.DefaultExt = ".xlsx";
            chooseOutputLocation.Filter = "Excel Workbook|*.xlsx";
            chooseOutputLocation.ShowDialog();
            txtFileOutput.Text = chooseOutputLocation.FileName;
            txtFileOutput.Enabled = false;
        }

        /*
         * SQL strings could be read from external files for neater
         * code but then the .sql files would have to be packaged
         * with and accessible to the application
         */

        private void setStrings() {
            /*
             *Updates query strings with current data from form
             */
            packagingUsedQuery = @"use BatchMetrics;

declare
@IPC_ID int = 0,
@StartDate dateTime = convert(datetime, '" + dtpStartDate.Value.Date.ToString("yyyy-MM-dd") + @"'),
@EndDate dateTime = convert(datetime, '" + dtpEndDate.Value.Date.ToString("yyyy-MM-dd") + @"');

select --ss.Description Status,  
bh.ScheduledStartDate 'Mfg. Date' ,
bh.Plant, 
bh.BatchNumber 'Mfg.#',    
u.Name Dryer,  
cust.Name Customer,  
ipc.IPC 'ERP Code', 
CASE WHEN u.Name LIKE '%5%' THEN bhp.Description ELSE ipc.Description END AS Description, 
case when Invoice_ERP_Document is not null  and Invoice_ERP_Document not like 'ORACLE%' THEN Invoice_ERP_Document ELSE
	case co.Invoice_ERP_ID when -1 then 'PENDING' else STR(Invoice_ERP_ID) end END as Invoice,
co.Invoice_Total,
co.GL_Date as 'filterDate'

from
	BatchHeader bh inner join BatchSub bs on bs.BatchHID = bh.BatchHID 
	join CustomerOrderMOLink mlink on mlink.BatchHid = bh.BatchHID 
	join CustomerOrderLine col on mlink.CustomerOrderLineID = col.LineID 
	join CustomerOrder co on col.CustomerOrderID = co.CustomerOrderID 
	join Customer cust on cust.CustomerID = co.CustomerID 
	left join IPCList ipc on ipc.IPC_ID = bh.ProductID 
	left join  BatchHeaderExtension bhe on bh.BatchHID = bhe.BatchHID 
	left join Units u on bs.UnitID = u.ID 
	left join IPCList_Extensions ie on ipc.IPC_ID=ie.IPC_ID 
	join StatusSub ss on bs.StatusID = ss.StatusID 
	left join dbo.BatchHeaderProduct AS bhp ON bhp.BatchHID = bh.BatchHID
    left join (select distinct [Key],[Value] from [Parameters] where [Type]=11 and Name ='Afterburner') params on ipc.IPC_ID = params.[Key]
	--join previousQuery on co.Invoice_ERP_Document = previousQuery.Invoice
	
where 
	bs.SubLevel = 1
	and Invoice_ERP_Document NOT IN(select
case when Invoice_ERP_Document is not null  and Invoice_ERP_Document not like 'ORACLE%' THEN Invoice_ERP_Document ELSE
	case co.Invoice_ERP_ID when -1 then 'PENDING' else STR(Invoice_ERP_ID) end END as Invoice

from 
	BatchHeader bh inner join BatchSub bs on bs.BatchHID = bh.BatchHID 
	join CustomerOrderMOLink mlink on mlink.BatchHid = bh.BatchHID 
	join CustomerOrderLine col on mlink.CustomerOrderLineID = col.LineID 
	join CustomerOrder co on col.CustomerOrderID = co.CustomerOrderID 
	join Customer cust on cust.CustomerID = co.CustomerID 
	left join IPCList ipc on ipc.IPC_ID = bh.ProductID 
	left join  BatchHeaderExtension bhe on bh.BatchHID = bhe.BatchHID 
	left join Units u on bs.UnitID = u.ID 
	left join IPCList_Extensions ie on ipc.IPC_ID=ie.IPC_ID 
	join StatusSub ss on bs.StatusID = ss.StatusID 
	left join dbo.BatchHeaderProduct AS bhp ON bhp.BatchHID = bh.BatchHID
    left join (select distinct [Key],[Value] from [Parameters] where [Type]=11 and Name ='Afterburner') params on ipc.IPC_ID = params.[Key]
	/*join Recipe r on r.IPCID = ipc.IPC_ID
	join RecipeSub rs on rs.RecipeID = r.ID
	join RecipeOperation ro on ro.SubID = rs.SubID*/
	join BatchLine bl ON bl.SubID = bs.SubID
	join BatchLineIPC blipc ON blipc.LineID = bl.LineID

where bs.SubLevel = 1
	and Invoice_ERP_Document IS NOT NULL
	and cust.CustomerID != 42629
	--and cust.CustomerID = 42629
	and col.LineNumber = 1 
	and bs.StatusID in (5)
	and (@IPC_ID = 0 or bh.ProductID = @IPC_ID)	
	and co.GL_Date between @StartDate and dbo.fn_Date_AlmostNextDay(@EndDate)
	and co.Invoice_ERP_Document is not null
	and blipc.IPCID in 
			(select ipc.IPC_ID 
			from IPCList ipc join IPCList_CustomerLink ipcl on ipcl.IPC_ID = ipc.IPC_ID 
			where ipc.ipc like '%-0000'	and ipc.MaterialTypeID IN (3,73) and ipcl.CustomerID = 9625 and ipc.IPC_ID not in (36189, 36190, 511848))

group by bh.BatchNumber, bh.ScheduledStartDate, bh.Plant, u.Name, cust.Name, ipc.IPC, ipc.Description, bhp.Description, co.Invoice_ERP_ID, co.Invoice_ERP_Document, co.Invoice_Total, co.GL_Date)
	and cust.CustomerID != 42629
	and col.LineNumber = 1 
	and bs.StatusID in (5)
	and (@IPC_ID = 0 or bh.ProductID = @IPC_ID)	
	and co.GL_Date between @StartDate and dbo.fn_Date_AlmostNextDay(@EndDate)
	and co.Invoice_ERP_Document is not null
	and ipc.IPC not like '%-9999'
	and ipc.IPC_ID in (select ipc.IPC_ID
							from IPCList ipc join FillRequirements fr on fr.ObjectID = ipc.IPC_ID
							where fr.ContainerTypeID in (select ipc.IPC_ID
														from IPCList ipc join IPCList_CustomerLink ipcl on ipcl.IPC_ID = ipc.IPC_ID
														where ipcl.CustomerID = 9625
														and ipc.MaterialTypeID = 71)
									or fr.PalletTypeID in (select ipc.IPC_ID
														from IPCList ipc join IPCList_CustomerLink ipcl on ipcl.IPC_ID = ipc.IPC_ID
														where ipcl.CustomerID = 9625
														and ipc.MaterialTypeID = 71)
									or fr.LinersCustomer in (select ipc.IPC_ID
														from IPCList ipc join IPCList_CustomerLink ipcl on ipcl.IPC_ID = ipc.IPC_ID
														where ipcl.CustomerID = 9625
														and ipc.MaterialTypeID = 71))


group by bh.BatchNumber, bh.ScheduledStartDate, bh.Plant, u.Name, cust.Name, ipc.IPC, ipc.Description, bhp.Description, co.Invoice_ERP_ID, co.Invoice_ERP_Document, co.Invoice_Total, co.GL_Date

order by Customer asc, 'Mfg. Date' desc";
            rawsUsedMcQuery = @"use BatchMetrics

declare
@IPC_ID int = 0,
@StartDate dateTime = convert(datetime, '" + dtpStartDate.Value.Date.ToString("yyyy-MM-dd") + @"'),
@EndDate dateTime = convert(datetime, '" + dtpEndDate.Value.Date.ToString("yyyy-MM-dd") + @"');
select
bh.ScheduledStartDate 'Mfg. Date' ,
bh.Plant, 
bh.BatchNumber 'Mfg.#',    
u.Name Dryer,  
cust.Name Customer,  
ipc.IPC 'ERP Code',  
CASE WHEN u.Name LIKE '%5%' THEN bhp.Description ELSE ipc.Description END AS Description, 
case when Invoice_ERP_Document is not null  and Invoice_ERP_Document not like 'ORACLE%' THEN Invoice_ERP_Document ELSE
	case co.Invoice_ERP_ID when -1 then 'PENDING' else STR(Invoice_ERP_ID) end END as Invoice,
co.Invoice_Total,
co.GL_Date as 'filterDate'

from 
	BatchHeader bh inner join BatchSub bs on bs.BatchHID = bh.BatchHID 
	join CustomerOrderMOLink mlink on mlink.BatchHid = bh.BatchHID 
	join CustomerOrderLine col on mlink.CustomerOrderLineID = col.LineID 
	join CustomerOrder co on col.CustomerOrderID = co.CustomerOrderID 
	join Customer cust on cust.CustomerID = co.CustomerID 
	left join IPCList ipc on ipc.IPC_ID = bh.ProductID 
	left join  BatchHeaderExtension bhe on bh.BatchHID = bhe.BatchHID 
	left join Units u on bs.UnitID = u.ID 
	left join IPCList_Extensions ie on ipc.IPC_ID=ie.IPC_ID 
	join StatusSub ss on bs.StatusID = ss.StatusID 
	left join dbo.BatchHeaderProduct AS bhp ON bhp.BatchHID = bh.BatchHID
    left join (select distinct [Key],[Value] from [Parameters] where [Type]=11 and Name ='Afterburner') params on ipc.IPC_ID = params.[Key]
	join Recipe r on r.IPCID = ipc.IPC_ID
	join RecipeSub rs on rs.RecipeID = r.ID
	join RecipeOperation ro on ro.SubID = rs.SubID

where bs.SubLevel = 1
	and col.LineNumber = 1 
	and bs.StatusID in (5)
	and (@IPC_ID = 0 or bh.ProductID = @IPC_ID)	
	and co.GL_Date between @StartDate and dbo.fn_Date_AlmostNextDay(@EndDate)
	and co.Invoice_ERP_Document is not null
	AND bh.ProductID IN(" + cat1cat2Switch + @")

group by bh.BatchNumber, bh.ScheduledStartDate, bh.Plant, u.Name, cust.Name, ipc.IPC, ipc.Description, bhp.Description, co.Invoice_ERP_ID, co.Invoice_ERP_Document, co.Invoice_Total, co.GL_Date

order by Customer ASC";
            monthlyInvoicesQuery = @"DECLARE
@StartDate dateTime = convert(datetime, '" + dtpStartDate.Value.Date.ToString("yyyy-MM-dd") + @"'),
@EndDate dateTime = convert(datetime, '" + dtpEndDate.Value.Date.ToString("yyyy-MM-dd") + @"')

SELECT 
	FilterDate AS 'Date', 
	Str(CustomerID) + '--' + Customer As Customer, 
	Invoice AS 'Inv no/Adjustment no', 
	Invoice_Total AS 'Amount'
FROM [BatchMetrics].[SprayTek].[tvf_CO_Batch_Dashboard] (
   /*CustomerID*/DEFAULT,/*IPC_ID*/DEFAULT,/*StartDate*/@StartDate,/*EndDate*/@EndDate,/*iStatus*/DEFAULT,/*sPlant*/DEFAULT
  ,/*iUnitID*/DEFAULT,/*customer*/DEFAULT,/*product*/DEFAULT)
UNION
SELECT 
	bm.DatePosted As Date,
	Str(bm.CustomerID) + '--' + Customer.Name As Customer,
	bm.InvoiceNumber AS 'Inv no/Adjustment no',
	bm.TotalAmount AS 'Amount'
FROM [BatchMetrics].[SprayTek].[Intacct_ManualInvoice] bm
JOIN Customer ON bm.CustomerID = Customer.CustomerID
WHERE DatePosted between @StartDate and @EndDate
	AND bm.TotalAmount <> 0
ORDER BY Invoice ASC";
            cat2SwitchQuery = @"Select ipc.IPC_ID
From IPCList ipc
where ipc.MaterialTypeID = 65
	and ipc.StatusID = 1
	and ipc.IPC_ID in (Select DISTINCT bh.ProductID
						from IPCList ipc join BatchHeader bh on bh.ProductID = ipc.IPC_ID
							join BatchSub bs on bs.BatchHID = bh.BatchHID
							join BatchLine bl on bl.SubID = bs.SubID
							join BatchLineIPC blipc ON blipc.LineID = bl.LineID
							JOIN CustomerOrderMOLink mlink ON mlink.BatchHid = bh.BatchHID 
							JOIN CustomerOrderLine col ON mlink.CustomerOrderLineID = col.LineID 
							JOIN CustomerOrder co ON col.CustomerOrderID = co.CustomerOrderID 
						where ipc.MaterialTypeID = 65
							and ipc.StatusID = 1
							and co.GL_Date BETWEEN @StartDate AND @EndDate
							and blipc.IPCID in (SELECT
														ipc.IPC_ID
												FROM
													IPCList ipc
													JOIN
														IPCList_CustomerLink ipcl
														ON
														ipcl.IPC_ID = ipc.IPC_ID
													JOIN
														IPCList_Extensions ipce
														ON
														ipce.IPC_ID = IPC.IPC_ID
												WHERE
													ipc.IPC LIKE '%-0000'
													AND
														ipc.MaterialTypeID IN (3,73)
													AND
														ipcl.CustomerID = 9625
													AND
														blipc.IPCID NOT IN (36189, 36190, 25535,25534,36171,511848)))
	-- below returns list of FG that use at least one ST category 1 raw material
	and ipc.IPC_ID not in (Select DISTINCT bh.ProductID
						from BatchHeader bh join BatchSub bs on bh.BatchHID = bs.BatchHID
							join BatchLine bl on bl.SubID = bs.SubID
							join BatchLineIPC blipc on blipc.LineID = bl.LineID
							join IPCList ipc ON ipc.IPC_ID = bh.ProductID
							JOIN CustomerOrderMOLink mlink ON mlink.BatchHid = bh.BatchHID 
							JOIN CustomerOrderLine col ON mlink.CustomerOrderLineID = col.LineID 
							JOIN CustomerOrder co ON col.CustomerOrderID = co.CustomerOrderID 
						where ipc.MaterialTypeID = 65
							and ipc.StatusID = 1
							and co.GL_Date  BETWEEN @StartDate AND @EndDate
							and blipc.IPCID in (SELECT
														ipc.IPC_ID
												FROM
													IPCList ipc
													JOIN
														IPCList_CustomerLink ipcl
														ON
														ipcl.IPC_ID = ipc.IPC_ID
													JOIN
														IPCList_Extensions ipce
														ON
														ipce.IPC_ID = IPC.IPC_ID
												WHERE
													ipc.IPC LIKE '%-0000'
													AND
														ipc.MaterialTypeID IN (3,73)
													AND
														ipcl.CustomerID = 9625
													AND
														blipc.IPCID NOT IN (36189, 36190, 25535,25534,36171,511848)
													AND ipce.[Category 1] = 1))";
            cat1SwitchQuery = @"Select ipc.IPC_ID
From IPCList ipc
where ipc.MaterialTypeID = 65
	and ipc.StatusID = 1
	and ipc.IPC_ID in (Select ipc.IPC_ID
						from BatchHeader bh join BatchSub bs on bh.BatchHID = bs.BatchHID
							join BatchLine bl on bl.SubID = bs.SubID
							join BatchLineIPC blipc on blipc.LineID = bl.LineID
							join IPCList ipc ON ipc.IPC_ID = bh.ProductID
							JOIN CustomerOrderMOLink mlink ON mlink.BatchHid = bh.BatchHID 
							JOIN CustomerOrderLine col ON mlink.CustomerOrderLineID = col.LineID 
							JOIN CustomerOrder co ON col.CustomerOrderID = co.CustomerOrderID 
						where ipc.MaterialTypeID = 65
							and ipc.StatusID = 1
							and co.GL_Date  BETWEEN @StartDate AND @EndDate
							and blipc.IPCID in (SELECT
														ipc.IPC_ID
												FROM
													IPCList ipc
													JOIN
														IPCList_CustomerLink ipcl
														ON
														ipcl.IPC_ID = ipc.IPC_ID
													JOIN
														IPCList_Extensions ipce
														ON
														ipce.IPC_ID = IPC.IPC_ID
												WHERE
													ipc.IPC LIKE '%-0000'
													AND
														ipc.MaterialTypeID IN (3,73)
													AND
														ipcl.CustomerID = 9625
													AND
														ipc.IPC_ID NOT IN (36189, 36190, 25535,25534,36171,511848)
													AND ipce.[Category 1] = 1))";
            cat4FullDetailQuery = @"select 
	MfgDate,
	Plant,
	[Mfg.#],
	Dryer,
	Customer,
	[ERP Code],
	Description,
	Invoice,
	Invoice_Total,
	filterDate

                      from [BatchMetrics].[SprayTek].[tvf_CO_Batch_Dashboard] (
                           /*CustomerID*/DEFAULT,/*IPC_ID*/DEFAULT,/*StartDate*/'" + dtpStartDate.Value.Date.ToString() + "',/*EndDate*/'" + dtpEndDate.Value.Date.ToString() + @"',/*iStatus*/DEFAULT,/*sPlant*/DEFAULT
                          ,/*iUnitID*/DEFAULT,/*customer*/DEFAULT,/*product*/DEFAULT)
WHERE
	Invoice IN (" + cat4Numbers + ")";
            manualInvoicesQuery = @"	SELECT 
		bm.DatePosted AS 'Date',
		Customer.Name AS Customer,
		bm.InvoiceNumber AS 'Inv no/Adjustment no',
		bm.TotalAmount AS 'Amount'
	FROM [BatchMetrics].[SprayTek].[Intacct_ManualInvoice] AS bm
		JOIN 
			Customer ON bm.CustomerID = Customer.CustomerID
	WHERE 
		DatePosted BETWEEN '" + dtpStartDate.Value.Date.ToString() + "' AND '" + dtpEndDate.Value.Date.ToString() + @"'
		AND bm.TotalAmount <> 0";
        }

        private void dtpStartDate_ValueChanged(object sender, EventArgs e)
        {
            setStrings();
        }

        private void dtpEndDate_ValueChanged(object sender, EventArgs e)
        {
            setStrings();
        }

        private void onClose(Object sender, EventArgs e) {
            /*
             *Overrides existing onClose method to make sure that
             *the SQL connection is closed when the application closes
             *
             *SQL connection may not close if application is closed in 
             *unorthodox ways
             */

            connection.Close();
        }
    }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;
        private DataSet time = null;
        private DataSet expense = null;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public int FldClient { get; set; }

        public int FldMatter { get; set; }

        public string origAtty = "";

        public string date = "";

        private bool attySelected = false;

        private bool dateSelected = false;

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

            string TkprIndex;
            cbFrom.ClearItems();
            string SQLTkpr = "select empinitials + case when len(empinitials)=1 then '     ' when len(empinitials)=2 then '     ' when len(empinitials)=3 then '   ' else '  ' end +  empname as employee from employee where  ( empsysnbr in (select morigatty from matorigatty) ) order by empinitials";
            DataSet myRSTkpr = _jurisUtility.RecordsetFromSQL(SQLTkpr);

            if (myRSTkpr.Tables[0].Rows.Count == 0)
                cbFrom.SelectedIndex = 0;
            else
            {
                foreach (DataTable table in myRSTkpr.Tables)
                {

                    foreach (DataRow dr in table.Rows)
                    {
                        TkprIndex = dr["employee"].ToString();
                        cbFrom.Items.Add(TkprIndex);
                    }
                }

            }

            DateTime result = DateTime.Today.Subtract(TimeSpan.FromDays(1));
            dateTimePicker1.Value = result;

        }



        #endregion

        #region Private methods

        private void DoDaFix()

        {
            

            if (attySelected && dateSelected)
            {
                DialogResult result = MessageBox.Show("You are about to remove all unbilled time and expense entries" + "\r\n" +
                    "for Originating Timekeeper: " + origAtty + " that occurred on or before" + "\r\n" +
                    "the following date: " + date + ". Are you sure?", "Confirmation Dialog", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == System.Windows.Forms.DialogResult.Yes)


                {
                    

                    string SQL = @"Select cast(count(*) as Varchar(10)) as PB from prebillmatter where pbmmatter in (select uematter from  unbilledexpense where uedate<='" + date + "' and uematter in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "'))" +
                        " or pbmmatter in (select utmatter from unbilledtime where utdate<='" + date + "' and utmatter in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "'))";
                    DataSet myRSPB = _jurisUtility.RecordsetFromSQL(SQL);
                    getPostReportDataSet();
                    foreach (DataTable table in myRSPB.Tables)
                    {

                        foreach (DataRow dr in table.Rows)
                        {
                           string PBIndex = dr["PB"].ToString();


                            if (PBIndex == "0")

                            {     // Enter your SQL code here
                                  // To run a T-SQL statement with no results, int RecordsAffected = _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                  // To get an ADODB.Recordset, ADODB.Recordset myRS = _jurisUtility.RecordsetFromSQL(SQL);

                                

                                UpdateStatus("Updating Time Detail Dist.", 0, 15);
                                SQL = "Delete from timedetaildist where tddid in (select utid from unbilledtime where utdate<='" + date + "' and utmatter in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "'))";

                             _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                                UpdateStatus("Updating Timekeeper Diary.", 1, 15);
                                SQL = "Delete from tkprdiary from timebatchdetail where tkprdiary.tdbatch=tbdbatch and tbdrecnbr=tdrecnbr and tbdid in (select utid from unbilledtime where utdate<='" + date + "' and utmatter in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "'))";
                                _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                                UpdateStatus("Updating Time Batch Detail.", 2,15);

                                SQL = "Delete from timebatchdetail where tbdid in (select utid from unbilledtime where utdate<='" + date + "' and utmatter in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "'))";
                              

                                _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                                UpdateStatus("Updating Fee Sum by Period.", 3, 15);

                                SQL = @"update feesumbyprd set fspworkedhrsentered=fspworkedhrsentered + hrs, 
                                    fspbilhrsentered=fspbilhrsentered + Bhrs, fspnonbilhrsentered=fspnonbilhrsentered + uhrs,
                                    fspfeeenteredactualvalue=fspfeeenteredactualvalue + amt,
                                            fspfeeenteredstdvalue=fspfeeenteredstdvalue+stdamt 
                                    from (select utmatter, uttkpr, isnull(uttaskcd,'') as taskcd, isnull(utactivitycd,'') as activitycd,
month(utdate) as utprdnbr,year(utdate) as utprdyear, sum(utactualhrswrk) * -1 as Hrs, sum(case when utbillableflg='Y' then utactualhrswrk else 0 end)  * -1 as Bhrs, 
                                sum(case when utbillableflg='N' then utactualhrswrk else 0 end) * -1 as UHrs, sum(utamount) * -1 as Amt, sum(utamtatstdrate) * -1 as StdAmt
                                from unbilledtime
                                where utdate<='" + date + "' and utmatter in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "')  group by utmatter, uttkpr, isnull(uttaskcd,''), isnull(utactivitycd,'') ,month(utdate), year(utdate)) UT   where utmatter=fspmatter and uttkpr=fsptkpr and utprdnbr=fspprdnbr and utprdyear=fspprdyear and isnull(fsptaskcd,'')=taskcd and isnull(fspactivitycd,'')=activitycd and fspworkedhrsentered<>hrs";

                                _jurisUtility.ExecuteNonQueryCommand(0, SQL);


                                UpdateStatus("Updating Unbilled Time.", 4, 15);

                                SQL = "Delete from unbilledtime where utdate<='" + date + "' and utmatter in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "')";

                                _jurisUtility.ExecuteNonQueryCommand(0, SQL);


                                UpdateStatus("Updating Expense Detail Dist.", 5, 15);
                                SQL = "Delete from expdetaildist where eddid in (select ueid from unbilledexpense where uedate<='" + date + "' and uematter in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "'))";

                                _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                                UpdateStatus("Updating Expense Batch Detail.", 5, 15);

                                SQL = "Delete from expbatchdetail where ebdid in (select ueid from unbilledexpense where  uedate<='" + date + "' and uematter in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "'))";
                                
                                
                                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                UpdateStatus("Updating Unbilled Expense.", 6, 15); 

                                SQL = "Delete from unbilledexpense where uedate<='" + date + "' and uematter in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "')";

                                _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                                UpdateStatus("Updating Time Batch Detail.", 7, 15);

                                SQL = "Delete from timebatchdetail where tbdid in (select tbdid from timeentrylink where entryid in (select entryid from timeentry where  entrydate<='" + date + "' and mattersysnbr in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "') and entrystatus<=7))";


                                _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                                UpdateStatus("Updating Time Entry Table.", 8, 15);

                                SQL = "Delete from timeentry where entrydate<='" + date + "' and mattersysnbr in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "') and entrystatus<=7";

                                _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                                SQL = "Delete from timeentrylink where entryid not in (select entryid from timeentry)";
                                _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                                UpdateStatus("Updating Expense Batch Detail.", 9, 15);

                                SQL = "Delete from expbatchdetail where ebdid in (select ebdid from expenseentrylink where entryid in (select entryid from expenseentry where entrydate<='" + date + "' and mattersysnbr in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "') and entrystatus<=7))";


                                _jurisUtility.ExecuteNonQueryCommand(0, SQL);


                                UpdateStatus("Updating Expense Entry Table.", 10, 15);

                                SQL = "Delete from expenseentry where entrydate<='" + date + "' and mattersysnbr in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "') and entrystatus<=7";

                                _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                                SQL = "Delete from expenseentrylink where entryid not in (select entryid from expenseentry)";
                                _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                                UpdateStatus("Updating Exp Sum by Period.", 11, 15);

                                SQL = @"update expsumbyprd set espentered=espentered + amt 
                                    from (select uematter, ueexpcd,ueprdnbr, ueprdyear, sum(ueamount) * -1 as Amt
                                from unbilledexpense
                                where uedate<='" + date + "' and uematter in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + @"')
                                group by uematter, ueexpcd,ueprdnbr, ueprdyear) UT
                                where uematter=espmatter and ueexpcd=espexpcd and ueprdnbr=espprdnbr and ueprdyear=espprdyear  ";

                                _jurisUtility.ExecuteNonQueryCommand(0, SQL);



                                UpdateStatus("Updating Fee Sum ITD Table.", 12, 15);

                                SQL = "update feesumitd set fsicurunbilhrs=ut, fsicurunbilbal=ua from (select utmatter, sum(case when utbillableflg='T' then utactualhrswrk else 0 end) as UT, sum(utamount) as UA from unbilledtime group by utmatter) UT where fsimatter=utmatter ";
                                _jurisUtility.ExecuteNonQueryCommand(0, SQL);


                                UpdateStatus("Updating Fee Sum ITD Table.", 14, 15);

                                SQL = "update expsumitd set esicurunbilbal=ua  from (select uematter, sum(ueamount) as ua from unbilledexpense group by uematter) UE where uematter=esimatter";
                                _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                                UpdateStatus("Time and Expense Entries dated on and before  " + date + " for Originating Timekeeper: " + origAtty + " have been deleted.",5, 5);
                                DialogResult r = MessageBox.Show("The process is complete! Would you like to view the deleted items in a report?", "Process Complete!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                if (r == DialogResult.Yes)
                                    runPostReport();
                            }
                            else
                            {
                                MessageBox.Show("There are open prebills for matters in selection. Delete prebills before proceeding", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                UpdateStatus("Application Cancelled.  Open Prebills Exist.", 0, 4);

                            }

                        }
                    }
                }
            }
            else
                MessageBox.Show("Please select both a date and an originting attorney", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
       
        }

        private void getPostReportDataSet()
        {
            string SQLtime = "Select dbo.jfn_formatclientcode(clicode) as Client, Clireportingname, matcode, matreportingname, " +
"utdate, empinitials as WorkTkpr, utamount, utactualhrswrk, utnarrative, 'UnBilledTime' as EntryType " +
"from unbilledtime " +
"inner join matter on utmatter=matsysnbr " +
"inner join client on matclinbr=clisysnbr " +
"inner join employee on empsysnbr=uttkpr " +
"where utdate<='" + date + "' and utmatter in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "') " +
"UNION ALL " +
"Select dbo.jfn_formatclientcode(clicode) as Client, Clireportingname, matcode, matreportingname, " +
"entrydate, empinitials as WorkTkpr, amount, actualhourswork, narrative, 'TimeEntry'  as EntryType " +
"from timeentry " +
"inner join matter on mattersysnbr=matsysnbr " +
"inner join client on matclinbr=clisysnbr " +
"inner join employee on empsysnbr=timekeepersysnbr " +
"where entrystatus<=7 and entrydate<='" + date + "' and mattersysnbr in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "') and entrystatus<=6"
;
            string SQLexpense = "Select dbo.jfn_formatclientcode(clicode) as Client, Clireportingname, matcode, matreportingname, " +
"uedate, ueexpcd as ExpCode, ueamount,  uenarrative, 'UnBilledExpense'  as EntryType " +
"from unbilledexpense  " +
"inner join matter on uematter=matsysnbr " +
"inner join client on matclinbr=clisysnbr " +
"where uedate<='" + date + "' and uematter in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "') " +
"UNION ALL " +
"Select dbo.jfn_formatclientcode(clicode) as Client, Clireportingname, matcode, matreportingname,  " +
"entrydate, expensecode, amount, narrative, 'ExpenseEntry'  as EntryType " +
"from expenseentry " +
"inner join matter on mattersysnbr=matsysnbr " +
"inner join client on matclinbr=clisysnbr " +
"where entrystatus<=7 and entrydate<='" + date + "' and mattersysnbr in (select morigmat from matorigatty inner join employee on empsysnbr=morigatty where empinitials='" + origAtty + "') and entrystatus<=6"
;

            time = _jurisUtility.RecordsetFromSQL(SQLtime);
            expense = _jurisUtility.RecordsetFromSQL(SQLexpense);
        }


        private void runPostReport()
        {
                ReportDisplay rpds = new ReportDisplay(time, expense);
                rpds.Show();
        }





        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }




        private void cbFrom_SelectedIndexChanged(object sender, EventArgs e)
        {
           
            origAtty = cbFrom.Text;
            origAtty = origAtty.Split(' ')[0];
            attySelected = true;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime result = dateTimePicker1.Value;
            date = result.ToString("yyyy-MM-dd");
            dateSelected = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            runPostReport();
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CovertDataAirPollution
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
        }
        BackgroundWorker _bgWorker = new BackgroundWorker();
        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Completed");
        }
        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            #region copy
            int index = 0;
            foreach (var s in files)
            {
                string fileName = Path.GetFileName(s);
                if (fileName[0] != 'T' || fileName[1] != 'D')
                    continue;

                int year = int.Parse("20" + fileName.Substring(2, 2));
                int month = int.Parse(fileName.Substring(4, 2));
                string stationID = fileName.Substring(fileName.Length - 7, 3);

                if (this.label1.InvokeRequired)
                {
                    this.label1.BeginInvoke((MethodInvoker)delegate()
                    {
                        label1.Text = ("Processing " + fileName + (++index) + "/" + files.Length);
                    });
                }
                else
                {
                    label1.Text = ("Processing " + fileName + (++index) + "/" + files.Length);
                }


                COlist = new List<APData>();

                CH4list = new List<APData>();
                NmHClist = new List<APData>();
                THClist = new List<APData>();
                O3list = new List<APData>();
                PM10list = new List<APData>();
                SO2list = new List<APData>();
                NOxlist = new List<APData>();
                NOlist = new List<APData>();
                NO2list = new List<APData>();
                TotalAPIlist = new List<APData>();
                AmbTemplist = new List<APData>();
                Humiditylist = new List<APData>();
                Windlist = new List<APData>();

                //NmHC Hourly
                //THC Hourly
                LoadDataFromExcel(s, "THC Hourly", ref THClist, stationID, year, month);
                LoadDataFromExcel(s, "NmHC Hourly", ref NmHClist, stationID, year, month);
                LoadDataFromExcel(s, "CH4 Hourly", ref CH4list, stationID, year, month);
                LoadDataFromExcel(s, "CO Hourly", ref COlist, stationID, year, month);
                LoadDataFromExcel(s, "O3 Hourly", ref O3list, stationID, year, month);
                LoadDataFromExcel(s, "PM10 Hourly", ref PM10list, stationID, year, month);
                LoadDataFromExcel(s, "SO2 Hourly", ref SO2list, stationID, year, month);
                LoadDataFromExcel(s, "NOx Hourly", ref NOxlist, stationID, year, month);
                LoadDataFromExcel(s, "NO Hourly", ref NOlist, stationID, year, month);
                LoadDataFromExcel(s, "NO2 Hourly", ref NO2list, stationID, year, month);
                LoadDataFromExcel(s, "TotalAPI Hourly", ref TotalAPIlist, stationID, year, month);
                LoadDataFromExcel(s, "AmbientTemp Hourly", ref AmbTemplist, stationID, year, month);
                LoadDataFromExcel(s, "Humidity Hourly", ref Humiditylist, stationID, year, month);
                LoadDataFromExcelWind(s, "WindSpeed Wind Hourly", ref Windlist, stationID, year, month);


                List<APDataCombined> list = (from pm in PM10list
                                             let o3 = O3list.Where(x => x.Year == pm.Year && x.Month == pm.Month && x.Day == pm.Day && x.Hour == pm.Hour).FirstOrDefault()
                                             let co = COlist.Where(x => x.Year == pm.Year && x.Month == pm.Month && x.Day == pm.Day && x.Hour == pm.Hour).FirstOrDefault()
                                             let so = SO2list.Where(x => x.Year == pm.Year && x.Month == pm.Month && x.Day == pm.Day && x.Hour == pm.Hour).FirstOrDefault()
                                             let nox = NOxlist.Where(x => x.Year == pm.Year && x.Month == pm.Month && x.Day == pm.Day && x.Hour == pm.Hour).FirstOrDefault()
                                             let no = NOlist.Where(x => x.Year == pm.Year && x.Month == pm.Month && x.Day == pm.Day && x.Hour == pm.Hour).FirstOrDefault()
                                             let no2 = NO2list.Where(x => x.Year == pm.Year && x.Month == pm.Month && x.Day == pm.Day && x.Hour == pm.Hour).FirstOrDefault()
                                             let api = TotalAPIlist.Where(x => x.Year == pm.Year && x.Month == pm.Month && x.Day == pm.Day && x.Hour == pm.Hour).FirstOrDefault()
                                             let amb = AmbTemplist.Where(x => x.Year == pm.Year && x.Month == pm.Month && x.Day == pm.Day && x.Hour == pm.Hour).FirstOrDefault()
                                             let hum = Humiditylist.Where(x => x.Year == pm.Year && x.Month == pm.Month && x.Day == pm.Day && x.Hour == pm.Hour).FirstOrDefault()
                                             let win = Windlist.Where(x => x.Year == pm.Year && x.Month == pm.Month && x.Day == pm.Day && x.Hour == pm.Hour).FirstOrDefault()
                                             let ch4 = CH4list.Where(x => x.Year == pm.Year && x.Month == pm.Month && x.Day == pm.Day && x.Hour == pm.Hour).FirstOrDefault()
                                             let NmHC = NmHClist.Where(x => x.Year == pm.Year && x.Month == pm.Month && x.Day == pm.Day && x.Hour == pm.Hour).FirstOrDefault()
                                             let THC = THClist.Where(x => x.Year == pm.Year && x.Month == pm.Month && x.Day == pm.Day && x.Hour == pm.Hour).FirstOrDefault()
                                             select new APDataCombined
                                             {
                                                 stationID = pm.stationID,
                                                 Year = pm.Year,
                                                 Month = pm.Month,
                                                 Day = pm.Day,
                                                 Hour = pm.Hour,
                                                 COVal = co != null ? co.DataVal : null,
                                                 CH4Val = ch4 != null ? ch4.DataVal : null,
                                                 O3Val = o3 != null ? o3.DataVal : null,
                                                 NmHCVal = NmHC != null ? NmHC.DataVal : null,
                                                 THCVal = THC != null ? THC.DataVal : null,
                                                 PM10Val = pm != null ? pm.DataVal : null,
                                                 SO2Val = so != null ? so.DataVal : null,
                                                 NOxVal = nox != null ? nox.DataVal : null,
                                                 NOVal = no != null ? no.DataVal : null,
                                                 NO2Val = no2 != null ? no2.DataVal : null,
                                                 APIVal = api != null ? api.DataVal : null,
                                                 AmbTempVal = amb != null ? amb.DataVal : null,
                                                 HumidityVal = hum != null ? hum.DataVal : null,
                                                 WindDirVal = win != null ? win.DataVal2 : "",
                                                 WindSpeedVal = win != null ? win.DataVal : null
                                             }
                                             ).ToList();
                // CreateCSV(ref list, stationID, year, month);
                //CreateExcel(ref list);

            }
            CreateCSV(ref Dailylist, "Combined", 0, 0, true);
            //CreateExcel(ref Dailylist, true);

            #endregion
        }

        List<APData> COlist = new List<APData>();
        List<APData> CH4list = new List<APData>();
        List<APData> NmHClist = new List<APData>();
        List<APData> THClist = new List<APData>();
        List<APData> O3list = new List<APData>();
        List<APData> PM10list = new List<APData>();
        List<APData> SO2list = new List<APData>();
        List<APData> NOxlist = new List<APData>();
        List<APData> NOlist = new List<APData>();
        List<APData> NO2list = new List<APData>();
        List<APData> TotalAPIlist = new List<APData>();
        List<APData> AmbTemplist = new List<APData>();
        List<APData> Humiditylist = new List<APData>();
        List<APData> Windlist = new List<APData>();

        List<APDataCombined> Dailylist = new List<APDataCombined>();
        string[] files = null;


        private void button1_Click(object sender, EventArgs e)
        {
            Dailylist = new List<APDataCombined>();
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    // string[] files = Directory.GetFiles(fbd.SelectedPath);

                    files = Directory.GetFiles(fbd.SelectedPath, "*.*", SearchOption.AllDirectories);


                    label1.Text = ("Files found: " + files.Length.ToString());

                    _bgWorker.RunWorkerAsync();



                }
            }

        }


        private void LoadDataFromExcel(string fileLoc, string sheetname, ref List<APData> list, string stationID, int year, int month)
        {
            String sConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;" +
                "Data Source=" + fileLoc + ";" +
                "Extended Properties=Excel 8.0;";

            OleDbConnection objConn = new OleDbConnection(sConnectionString);

            try
            {

                // Open connection with the database.
                objConn.Open();

                // The code to follow uses a SQL SELECT command to display the data from the worksheet.

                // Create new OleDbCommand to return data from worksheet.
                OleDbCommand objCmdSelect = new OleDbCommand("Select * from [" + sheetname + "$]", objConn);

                // Create new OleDbDataAdapter that is used to build a DataSet
                // based on the preceding SQL SELECT statement.
                OleDbDataAdapter objAdapter1 = new OleDbDataAdapter();

                // Pass the Select command to the adapter.
                objAdapter1.SelectCommand = objCmdSelect;

                // Create new DataSet to hold information from the worksheet.
                DataSet objDataset1 = new DataSet();

                // Fill the DataSet with the information from the worksheet.
                objAdapter1.Fill(objDataset1, "XLData");

                // Bind data to DataGrid control.

                //  dataGridView1.DataSource = objDataset1.Tables[0].DefaultView;

                //List<ApplicationQualification> list = new List<ApplicationQualification>();
                var s = objDataset1.Tables[0].AsEnumerable();
                foreach (var a in s)
                {

                    //ApplicationQualification ap = new ApplicationQualification();

                    if (!string.IsNullOrEmpty(a.ItemArray[0].ToString().Trim()))
                    {
                        int tryInt;

                        if (int.TryParse(a.ItemArray[0].ToString().Trim(), out tryInt))
                        {
                            for (int i = 1; i <= 24; i++)
                            {
                                APData l = new APData();
                                l.Year = year;
                                l.Month = month;
                                l.Day = int.Parse(a.ItemArray[0].ToString().Trim());
                                l.Hour = i;
                                float val;
                                if (float.TryParse(a.ItemArray[i].ToString().Trim(), out val))
                                    l.DataVal = val;
                                else
                                    l.DataVal = null;

                                l.stationID = stationID;

                                list.Add(l);
                            }
                            int day = int.Parse(a.ItemArray[0].ToString().Trim());

                            APDataCombined cl = Dailylist.Where(x => x.Year == year && x.Month == month && x.Day == day && x.stationID == stationID).FirstOrDefault();
                            bool isNew = false;
                            if (cl == null)
                            {
                                cl = new APDataCombined();
                                cl.Year = year;
                                cl.Month = month;
                                cl.Day = day;
                                cl.stationID = stationID;
                                isNew = true;
                            }
                            float? nulableVal = null;
                            float val25;
                            if (float.TryParse(a.ItemArray[25].ToString().Trim(), out val25))
                                nulableVal = val25;

                            switch (sheetname)
                            {
                                case "CO Hourly": { cl.COVal = nulableVal; break; }
                                case "CH4 Hourly": { cl.CH4Val = nulableVal; break; }
                                case "NmHC Hourly": { cl.NmHCVal = nulableVal; break; }
                                case "THC Hourly": { cl.THCVal = nulableVal; break; }
                                case "O3 Hourly": { cl.O3Val = nulableVal; break; }
                                case "PM10 Hourly": { cl.PM10Val = nulableVal; break; }
                                case "SO2 Hourly": { cl.SO2Val = nulableVal; break; }
                                case "NOx Hourly": { cl.NOxVal = nulableVal; break; }
                                case "NO Hourly": { cl.NOVal = nulableVal; break; }
                                case "NO2 Hourly": { cl.NO2Val = nulableVal; break; }
                                case "TotalAPI Hourly": { cl.APIVal = nulableVal; break; }
                                case "AmbientTemp Hourly": { cl.AmbTempVal = nulableVal; break; }
                                case "Humidity Hourly": { cl.HumidityVal = nulableVal; break; }
                            }
                            if (isNew)
                                Dailylist.Add(cl);


                        }

                    }
                }


                // Clean up objects.
                objConn.Close();
            }
            catch (Exception ex)
            {

                objConn.Close();
                File.AppendAllText("log.txt", Environment.NewLine + "Exception: " + DateTime.Now.ToString() + "#sheetname: " + sheetname + "#stationID: " + stationID + "#year: " + year + "#month" + month + "#Exception Text: " + ex.Message);

            }

            //WriteToExcelApplicantQualification(list);

        }
        private void LoadDataFromExcelWind(string fileLoc, string sheetname, ref List<APData> list, string stationID, int year, int month)
        {
            String sConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;" +
                "Data Source=" + fileLoc + ";" +
                "Extended Properties=Excel 8.0;";

            OleDbConnection objConn = new OleDbConnection(sConnectionString);

            try
            {
                // Create connection string variable. Modify the "Data Source"
                // parameter as appropriate for your environment.


                // Create connection object by using the preceding connection string.

                // Open connection with the database.
                objConn.Open();

                // The code to follow uses a SQL SELECT command to display the data from the worksheet.

                // Create new OleDbCommand to return data from worksheet.
                OleDbCommand objCmdSelect = new OleDbCommand("Select * from [" + sheetname + "$]", objConn);

                // Create new OleDbDataAdapter that is used to build a DataSet
                // based on the preceding SQL SELECT statement.
                OleDbDataAdapter objAdapter1 = new OleDbDataAdapter();

                // Pass the Select command to the adapter.
                objAdapter1.SelectCommand = objCmdSelect;

                // Create new DataSet to hold information from the worksheet.
                DataSet objDataset1 = new DataSet();

                // Fill the DataSet with the information from the worksheet.
                objAdapter1.Fill(objDataset1, "XLData");

                // Bind data to DataGrid control.

                //  dataGridView1.DataSource = objDataset1.Tables[0].DefaultView;

                //List<ApplicationQualification> list = new List<ApplicationQualification>();
                var s = objDataset1.Tables[0].AsEnumerable();
                foreach (var a in s)
                {

                    //ApplicationQualification ap = new ApplicationQualification();

                    if (!string.IsNullOrEmpty(a.ItemArray[0].ToString().Trim()))
                    {
                        int tryInt;

                        if (int.TryParse(a.ItemArray[0].ToString().Trim(), out tryInt))
                        {
                            for (int i = 1, j = 1; i <= 24; i++, j++)
                            {
                                APData l = new APData();
                                l.Year = year;
                                l.Month = month;
                                l.Day = int.Parse(a.ItemArray[0].ToString().Trim());
                                l.Hour = i;

                                l.DataVal2 = a.ItemArray[j].ToString().Trim();
                                l.stationID = stationID;

                                float val;
                                if (float.TryParse(a.ItemArray[(++j)].ToString().Trim(), out val))
                                    l.DataVal = val;
                                else
                                    l.DataVal = null;

                                list.Add(l);
                            }
                            int day = int.Parse(a.ItemArray[0].ToString().Trim());

                            APDataCombined cl = Dailylist.Where(x => x.Year == year && x.Month == month && x.Day == day && x.stationID == stationID).FirstOrDefault();
                            bool isNew = false;

                            if (cl == null)
                            {
                                cl = new APDataCombined();
                                cl.Year = year;
                                cl.Month = month;
                                cl.Day = day;
                                cl.stationID = stationID;
                                isNew = true;
                            }
                            float val49;
                            if (float.TryParse(a.ItemArray[49].ToString().Trim(), out val49))
                                cl.WindSpeedVal = val49;
                            else
                                cl.WindSpeedVal = null;


                            if (isNew)
                                Dailylist.Add(cl);
                        }

                    }
                }


                // Clean up objects.
                objConn.Close();
            }
            catch (Exception ex)
            {
                File.AppendAllText("log.txt", Environment.NewLine + "Exception: " + DateTime.Now.ToString() + "#sheetname: " + sheetname + "#stationID: " + stationID + "#year: " + year + "#month" + month + "#Exception Text: " + ex.Message);
                objConn.Close();

            }

            //WriteToExcelApplicantQualification(list);

        }
        public void CreateExcel(ref List<APDataCombined> list, string stationID, int year, int month, bool isCombined = false)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;
            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = false;

                //Get a new workbook.
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                if (!isCombined)
                {
                    oSheet.Cells[1, 1] = "Year";
                    oSheet.Cells[1, 2] = "Month";
                    oSheet.Cells[1, 3] = "Day";
                    oSheet.Cells[1, 4] = "Hour";
                    oSheet.Cells[1, 5] = "CO";
                    oSheet.Cells[1, 6] = "O3";
                    oSheet.Cells[1, 7] = "PM10";
                    oSheet.Cells[1, 8] = "NOx";
                    oSheet.Cells[1, 9] = "NO2";
                    oSheet.Cells[1, 10] = "NO";
                    oSheet.Cells[1, 11] = "SO2";
                    oSheet.Cells[1, 12] = "Total API";
                    oSheet.Cells[1, 13] = "Ambient Temp";
                    oSheet.Cells[1, 14] = "Humidity";
                    oSheet.Cells[1, 15] = "Wind Dir";
                    oSheet.Cells[1, 16] = "Wind Speed";
                    oSheet.Cells[1, 17] = "StationID";
                }
                else
                {
                    oSheet.Cells[1, 1] = "Year";
                    oSheet.Cells[1, 2] = "Month";
                    oSheet.Cells[1, 3] = "Day";
                    oSheet.Cells[1, 4] = "CO";
                    oSheet.Cells[1, 5] = "O3";
                    oSheet.Cells[1, 6] = "PM10";
                    oSheet.Cells[1, 7] = "NOx";
                    oSheet.Cells[1, 8] = "NO2";
                    oSheet.Cells[1, 9] = "NO";
                    oSheet.Cells[1, 10] = "SO2";
                    oSheet.Cells[1, 11] = "Total API";
                    oSheet.Cells[1, 12] = "Ambient Temp";
                    oSheet.Cells[1, 13] = "Humidity";
                    oSheet.Cells[1, 14] = "Wind Speed";
                    oSheet.Cells[1, 15] = "StationID";
                }

                list = list.OrderBy(x => x.stationID).OrderBy(x => x.Year).ThenBy(x => x.Month).ThenBy(x => x.Day).ThenBy(x => x.Hour).ToList();

                int i = 2;
                foreach (APDataCombined d in list)
                {

                    if (!isCombined)
                    {
                        oSheet.Cells[i, 1] = d.Year;
                        oSheet.Cells[i, 2] = d.Month;
                        oSheet.Cells[i, 3] = d.Day;
                        oSheet.Cells[i, 4] = d.Hour;
                        oSheet.Cells[i, 5] = d.COVal;
                        oSheet.Cells[i, 6] = d.O3Val;
                        oSheet.Cells[i, 7] = d.PM10Val;
                        oSheet.Cells[i, 8] = d.NOxVal;
                        oSheet.Cells[i, 9] = d.NO2Val;
                        oSheet.Cells[i, 10] = d.NOVal;
                        oSheet.Cells[i, 11] = d.SO2Val;
                        oSheet.Cells[i, 12] = d.APIVal;
                        oSheet.Cells[i, 13] = d.AmbTempVal;
                        oSheet.Cells[i, 14] = d.HumidityVal;
                        oSheet.Cells[i, 15] = d.WindDirVal;
                        oSheet.Cells[i, 16] = d.WindSpeedVal;
                        oSheet.Cells[i, 17] = d.stationID;
                    }
                    else
                    {
                        oSheet.Cells[i, 1] = d.Year;
                        oSheet.Cells[i, 2] = d.Month;
                        oSheet.Cells[i, 3] = d.Day;
                        oSheet.Cells[i, 4] = d.COVal;
                        oSheet.Cells[i, 5] = d.O3Val;
                        oSheet.Cells[i, 6] = d.PM10Val;
                        oSheet.Cells[i, 7] = d.NOxVal;
                        oSheet.Cells[i, 8] = d.NO2Val;
                        oSheet.Cells[i, 9] = d.NOVal;
                        oSheet.Cells[i, 10] = d.SO2Val;
                        oSheet.Cells[i, 11] = d.APIVal;
                        oSheet.Cells[i, 12] = d.AmbTempVal;
                        oSheet.Cells[i, 13] = d.HumidityVal;
                        oSheet.Cells[i, 14] = d.WindSpeedVal;
                        oSheet.Cells[i, 15] = d.stationID;

                    }
                    i++;
                }


                oXL.Visible = false;
                oXL.UserControl = false;

                string root = @"D:\Output\" + stationID;
                // If directory does not exist, create it. 
                if (!Directory.Exists(root))
                {
                    Directory.CreateDirectory(root);
                }

                oWB.SaveAs("D:\\Output\\" + stationID + "\\" + stationID + "-" + year + "-" + (isCombined ? "Combined" : month.ToString()) + "-" + DateTime.Now.Ticks + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                oWB.Close();


            }
            catch (Exception ex)
            {
                File.AppendAllText("log.txt", Environment.NewLine + "Exception: " + DateTime.Now.ToString() + " #stationID: " + stationID + "#year: " + year + "#month" + month + "#Exception Text: " + ex.Message);
                MessageBox.Show(ex.Message);
            }
        }

        public void CreateCSV(ref List<APDataCombined> list, string stationID, int year, int month, bool isCombined = false)
        {
            string root = @"D:\OutputNewSingle\" + stationID;
            // If directory does not exist, create it. 
            if (!Directory.Exists(root))
            {
                Directory.CreateDirectory(root);
            }
            string path = "D:\\OutputNewSingle\\" + stationID + "\\" + stationID + "-" + year + "-" + (isCombined ? "Combined" : month.ToString()) + "-" + DateTime.Now.Ticks + ".csv";


            try
            {
                int i = 2;
                list = list.OrderBy(x => x.stationID).OrderBy(x => x.Year).ThenBy(x => x.Month).ThenBy(x => x.Day).ThenBy(x => x.Hour).ToList();
                using (var w = new StreamWriter(path))
                {
                    string line = "";


                    if (!isCombined)
                    {
                        line = "StationID,Year,Month,Day,Hour,CO,O3,PM10,NOx,NO2,NO,SO2,CH4,NmHC,THC,Total API,Ambient Temp,Humidity,Wind Dir,Wind Speed";
                    }
                    else
                    {
                        line = "StationID,Year,Month,Day,CO,O3,PM10,NOx,NO2,NO,SO2,CH4,NmHC,THC,Total API,Ambient Temp,Humidity,Wind Speed";
                    }
                    w.WriteLine(line);
                    w.Flush();

                }
                using (var w = new StreamWriter(path, true))
                {
                    foreach (var d in list)
                    {
                        string line = "";

                        if (!isCombined)
                        {
                            line = d.stationID + "," + d.Year + "," + d.Month + "," + d.Day + "," + d.Hour + "," + d.COVal + "," + d.O3Val + "," + d.PM10Val
                                + "," + d.NOxVal + "," + d.NO2Val + "," + d.NOVal + "," + d.SO2Val + "," + d.CH4Val + "," + d.NmHCVal + "," + d.THCVal + "," +
                                d.APIVal + "," + d.AmbTempVal + "," + d.HumidityVal + "," + d.WindDirVal + "," + d.WindSpeedVal;


                        }
                        else
                        {
                            line = d.stationID + "," + d.Year + "," + d.Month + "," + d.Day + "," + d.COVal + "," + d.O3Val + "," + d.PM10Val
                                + "," + d.NOxVal + "," + d.NO2Val + "," + d.NOVal + "," + d.SO2Val + "," + d.CH4Val + "," + d.NmHCVal + "," + d.THCVal + "," +
                                d.APIVal + "," + d.AmbTempVal + "," + d.HumidityVal + "," + d.WindSpeedVal;
                        }
                        i++;
                        w.WriteLine(line);
                        w.Flush();
                    }
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText("log.txt", Environment.NewLine + "Exception: " + DateTime.Now.ToString() + " #stationID: " + stationID + "#year: " + year + "#month" + month + "#Exception Text: " + ex.Message);
                MessageBox.Show(ex.Message);
            }
        }


    }
    public class APData
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public int Day { get; set; }
        public int Hour { get; set; }

        public float? DataVal { get; set; }
        public string DataVal2 { get; set; }
        public string stationID { get; set; }

    }
    public class APDataCombined
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public int Day { get; set; }
        public int Hour { get; set; }
        public string stationID { get; set; }

        public float? COVal { get; set; }
        public float? CH4Val { get; set; }
        public float? NmHCVal { get; set; }

        public float? THCVal { get; set; }


        public float? O3Val { get; set; }
        public float? PM10Val { get; set; }
        public float? SO2Val { get; set; }
        public float? NOxVal { get; set; }
        public float? NOVal { get; set; }
        public float? NO2Val { get; set; }
        public float? APIVal { get; set; }
        public float? AmbTempVal { get; set; }
        public float? HumidityVal { get; set; }

        public string WindDirVal { get; set; }
        public float? WindSpeedVal { get; set; }
    }
}

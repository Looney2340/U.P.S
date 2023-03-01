using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Newtonsoft;
using HtmlAgilityPack;
using System.Web;
using System.Security.Cryptography.X509Certificates;
using System.Diagnostics;

namespace WindowsFormsApplication1
{    public partial class Form1 : Form
    {
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        public class mainConfig
        {
            public string Building { get; set; }
            public string SLIC { get; set; }
        }
        public class drvrData
        {
            public string Route { get; set; }
            public string Drivername { get; set; }
            public string EmployeeNumber { get; set; }
            public string DelStops { get; set; }
            public string DelPkgs { get; set; }
            public string PUStops { get; set; }
            public string PUPkgs { get; set; }
            public string Miles { get; set; }
            public string SPM { get; set; }
            public string PlanDay { get; set; }
            public string PaidDay { get; set; }
            public string OvUn { get; set; }
            public string SPORH { get; set; }
            public string NDPPH { get; set; }
            public string AM { get; set; }
            public string PM { get; set; }
            public string NumDays { get; set; }
        }
        public class dateDriver
        {
            public string Date { get; set; }
            public string Driver { get; set; }
            public string DelStops { get; set; }
            public string DelPkgs { get; set; }
            public string PUStops { get; set; }
            public string PUPkgs { get; set; }
            public string Miles { get; set; }
            public string SPM { get; set; }
            public string PlanDay { get; set; }
            public string PaidDay { get; set; }
            public string OvUn { get; set; }
            public string SPORH { get; set; }
            public string NDPPH { get; set; }
            public string AM { get; set; }
            public string PM { get; set; }
            public string Route { get; set; }
            public string EmployeeID { get; set; }
        }
        public class thirtydaydriverDatawDate
        {
            public string driverName { get; set; }
            public string employeeID { get; set; }
            public string supGroup { get; set; }
            public string paycode { get; set; }
            public string delDate { get; set; }
            public string route { get; set; }
            public string planDay { get; set; }
            public string OvUn { get; set; }
            public string paidDay { get; set; }
            public string SPORH { get; set; }
            public string miles { get; set; }
            public string SPM { get; set; }
            public string delStops { get; set; }
            public string puStops { get; set; }
            public string delPkgs { get; set; }
            public string puPkgs { get; set; }
            public string paidSApercent { get; set; }
            public string ccsendAgains { get; set; }
            public string AM { get; set; }
            public string PM { get; set; }
            public string onroadHours { get; set; }
            public string NDPPH { get; set; }
            public string helperHours { get; set; }
        }
        public class thirtydaydriverData
        {
            public string driverName { get; set; }
            public string employeeID { get; set; }
            public string supGroup { get; set; }
            public string paycode { get; set; }
            public string route { get; set; }
            public string dayWork { get; set; }
            public string planDay { get; set; }
            public string OvUn { get; set; }
            public string paidDay { get; set; }
            public string SPORH { get; set; }
            public string miles { get; set; }
            public string SPM { get; set; }
            public string delStops { get; set; }
            public string puStops { get; set; }
            public string delPkgs { get; set; }
            public string puPkgs { get; set; }
            public string paidSApercent { get; set; }
            public string ccsendAgains { get; set; }
            public string AM { get; set; }
            public string PM { get; set; }
            public string onroadHours { get; set; }
            public string NDPPH { get; set; }
            public string helperHours { get; set; }
        }
        public class ORSS
        {
            public string AssignmentValdationCode { get; set; }
            public string BreakDuration { get; set; }
            public Boolean CanEdit { get; set; }
            public string ContractualLunch { get; set; }
            public string DispatchStatusCode { get; set; }
            public string DisplayBreakDuration { get; set; }
            public string DisplayStartTime { get; set; }
            public string DriverStatus { get; set; }
            public string DriverType { get; set; }
            public string DriverTypeCode { get; set; }
            public Boolean DynamicRoute { get; set; }
            public string EDDStatus { get; set; }
            public string ETAAdjMaxLimit { get; set; }
            public string EmployeeNumber { get; set; }
            public string FirstName { get; set;}
            public string HelperHours { get; set; }
            public string ID { get; set; }
            public string LastName { get; set; }
            public string LeaveTime { get; set; }
            public string LoadAreaName { get; set; }
            public string LoadPosition { get; set; }
            public string LoggedInUser { get; set; }
            public string LoginTime { get; set; }
            public string ManifestMaster { get; set; }
            public string ManifestSourceRole { get; set; }
            public string ManifestStatus { get; set; }
            public string Nickname { get; set; }
            public string OrganizationNumber { get; set; }
            public string PasVersion { get; set; }
            public string PayCode { get; set; }
            public string PerformanceHours { get; set; }
            public plannedTime PlannedTimecard { get; set; }    
            public string PreloadPASVersionNumber { get; set; }
            public string PreloadReferentNumber { get; set; }
            public string PreloadReferrant { get; set; }
            public string PreloadSortDate { get; set; }
            public string PreloadSortDateForDisplay { get; set; }
            public string PreloadSystemNumber { get; set; }
            public string PreloadUniqueIdentifier { get; set; }
            public string RouteName { get; set; }
            public string RouteStatusCode { get; set; }
            public string RouteType { get; set; }
            public string RouteUniqueIdentifier { get; set; }
            public string ServiceProviderType { get; set; }
            public string ServiceProviderTypeMean { get; set; }
            public string ServiceProviderTypeVal { get; set; }
            public string SortMaster { get; set; }
            public string SortServiceDate { get; set; }
            public string SortStartTime { get; set; }
            public string StartTime { get; set; }
            public string SupGroup { get; set; }
            public string SupervisorGroup { get; set; }
            public string TargetPaidHours { get; set; }
            public string abSorts { get; set; }
            public string preloadId { get; set; }
        }

        public class plannedTime
        {
            public string DriverType { get; set; }
            public string PayCode { get; set; }
            public string SupGroup { get; set; }
        }

        public void dateData()
        {
            Boolean detected = false;
            int placeHolder = 0;
            int driverLine = 0;
            int i = 0;
            listBox1.Items.Clear();
            try
            {
                WebClient wClient = new WebClient();
                var data = wClient.DownloadString($"https://pft.inside.ups.com/apps/DPSReports/SPperfomanceTrendMonthDtl.aspx?RTE={comboBox2.Text}&EMPID={employeeID.Text}&BLDG={bLDG.Text}&SLIC={sLIC.Text}");
                HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
                hDoc.LoadHtml( data );

                var tableQuery = from table in hDoc.DocumentNode.SelectNodes("//table").Cast<HtmlNode>()
                                 from row in table.SelectNodes("tr").Cast<HtmlNode>()
                                 from cell in row.SelectNodes("th|td").Cast<HtmlNode>()
                                 select new { Table = table.Id, CellText = cell.InnerText };

                List<dateDriver> drvrDate = new List<dateDriver>();

                foreach (var line in tableQuery)
                {
                    if (line.CellText.ToString() == "DelDate")
                    {
                        detected = true;
                        driverLine = placeHolder + 15;
                    }
                    else
                    {
                        if (detected == false)
                        {
                            placeHolder++;
                        }
                    }

                    if (detected == true)
                    {
                        if (line.CellText.Contains("10 Days"))
                        {
                            break;
                        }
                        if (line.CellText.Length == 5 || line.CellText.Length == 6) 
                        {
                            if (line.CellText.Trim().Contains("Day") || line.CellText.Trim().Contains("Days"))
                            {
                                break;
                            }
                        }
                        if (driverLine == placeHolder)
                        {
                            var driverA = new dateDriver() { Date = line.CellText.Trim() };
                            drvrDate.Add(driverA);
                        }

                        if (driverLine + 1 == placeHolder)
                        {
                            drvrDate[i].Driver = line.CellText.Trim();
                        }
                        if (driverLine +2 == placeHolder)
                        {
                            drvrDate[i].DelStops = line.CellText.Trim();
                        }
                        if (driverLine + 3 == placeHolder)
                        {
                            drvrDate[i].DelPkgs = line.CellText.Trim();
                        }
                        if (driverLine + 4 == placeHolder)
                        {
                            drvrDate[i].PUStops = line.CellText.Trim();
                        }
                        if (driverLine + 5 == placeHolder)
                        {
                            drvrDate[i].PUPkgs = line.CellText.Trim();
                        }
                        if (driverLine + 6 == placeHolder)
                        {
                            drvrDate[i].Miles = line.CellText.Trim();
                        }
                        if (driverLine + 7 == placeHolder)
                        {
                            drvrDate[i].SPM = line.CellText.Trim();
                        }
                        if (driverLine + 8 == placeHolder)
                        {
                            drvrDate[i].PlanDay = line.CellText.Trim();
                        }
                        if (driverLine + 9 == placeHolder)
                        {
                            drvrDate[i].PaidDay = line.CellText.Trim();
                        }
                        if (driverLine + 10 == placeHolder)
                        {
                            drvrDate[i].OvUn = line.CellText.Trim();
                        }
                        if (driverLine + 11 == placeHolder)
                        {
                            drvrDate[i].SPORH = line.CellText.Trim();
                        }
                        if (driverLine + 12 == placeHolder)
                        {
                            drvrDate[i].NDPPH = line.CellText.Trim();
                        }
                        if (driverLine + 13 == placeHolder)
                        {
                            drvrDate[i].AM = line.CellText.Trim();
                        }
                        if (driverLine + 14 == placeHolder)
                        {
                            drvrDate[i].PM = line.CellText.Trim();
                            drvrDate[i].Route = comboBox2.Text;
                            drvrDate[i].EmployeeID = employeeID.Text;
                            i++;
                            placeHolder++;
                            driverLine = placeHolder;
                        }
                        else
                        {
                            placeHolder++;
                        }
                    }

                }

                foreach (var lines in drvrDate)
                {
                    listBox1.Items.Add(lines.Date);
                }

                if (!File.Exists("driverdateData.json"))
                {
                    string jsonDrivers = JsonConvert.SerializeObject(drvrDate, Newtonsoft.Json.Formatting.Indented);
                    File.WriteAllText("driverdateData.json", jsonDrivers);
                }
                else
                {
                    File.Delete("driverdateData.json");
                    string jsonDrivers = JsonConvert.SerializeObject(drvrDate, Newtonsoft.Json.Formatting.Indented);
                    File.WriteAllText("driverdateData.json", jsonDrivers);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}", "ERROR!", buttons: MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void thirtydayData()
        {
            Boolean detected = false;
            int placeHolder = 0;
            int driverLine = 0;
            int i = 0;
            try
            {
                WebClient wClient = new WebClient();
                var data = wClient.DownloadString($"https://pft.inside.ups.com/apps/DPSReports/SPMonthAvg.aspx?BLDG={bLDG.Text}&SLIC={sLIC.Text}");
                HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
                hDoc.LoadHtml(data);


                var tableQuery = from table in hDoc.DocumentNode.SelectNodes("//table").Cast<HtmlNode>()
                                 from row in table.SelectNodes("tr").Cast<HtmlNode>()
                                 from cell in row.SelectNodes("th|td").Cast<HtmlNode>()
                                 select new { Table = table.Id, CellText = cell.InnerText };

                List<thirtydaydriverData> drivers = new List<thirtydaydriverData>();

                foreach (var line in tableQuery)
                {

                    if (line.CellText.ToString() == "Rte")
                    {
                        detected = true;
                        driverLine = placeHolder + 19;
                    }
                    else
                    {
                        if (detected == false)
                        {
                            placeHolder++;
                        }
                    }
                    if (detected == true)
                    {
                        if (driverLine == placeHolder)
                        {
                            var driverA = new thirtydaydriverData() { driverName = line.CellText.Trim() };
                            drivers.Add(driverA);
                        }
                        if (driverLine + 1 == placeHolder)
                        {
                            drivers[i].employeeID = line.CellText.Trim();
                        }
                        if (driverLine + 2 == placeHolder)
                        {
                            drivers[i].supGroup = line.CellText.Trim();
                        }
                        if (driverLine + 3 == placeHolder)
                        {
                            drivers[i].paycode = line.CellText.Trim();
                        }
                        if (driverLine + 4 == placeHolder)
                        {
                            drivers[i].route = line.CellText.Trim();
                        }
                        if (driverLine + 5 == placeHolder)
                        {
                            drivers[i].dayWork = line.CellText.Trim();
                        }
                        if (driverLine + 6 == placeHolder)
                        {
                            drivers[i].planDay = line.CellText.Trim();
                        }
                        if (driverLine + 7 == placeHolder)
                        {
                            drivers[i].OvUn = line.CellText.Trim();
                        }
                        if (driverLine + 8 == placeHolder)
                        {
                            drivers[i].paidDay = line.CellText.Trim();
                        }
                        if (driverLine + 9 == placeHolder)
                        {
                            drivers[i].SPORH = line.CellText.Trim();
                        }
                        if (driverLine + 10 == placeHolder)
                        {
                            drivers[i].miles = line.CellText.Trim();
                        }
                        if (driverLine + 11 == placeHolder)
                        {
                            drivers[i].SPM = line.CellText.Trim();
                        }
                        if (driverLine + 12 == placeHolder)
                        {
                            drivers[i].delStops = line.CellText.Trim();
                        }
                        if (driverLine + 13 == placeHolder)
                        {
                            drivers[i].puStops = line.CellText.Trim();
                        }
                        if (driverLine + 14 == placeHolder)
                        {
                            drivers[i].delPkgs = line.CellText.Trim();
                        }
                        if (driverLine + 15 == placeHolder)
                        {
                            drivers[i].puPkgs = line.CellText.Trim();
                        }
                        if (driverLine + 16 == placeHolder)
                        {
                            drivers[i].paidSApercent = line.CellText.Trim();
                        }
                        if (driverLine + 17 == placeHolder)
                        {
                            drivers[i].ccsendAgains = line.CellText.Trim();
                        }
                        if (driverLine + 18 == placeHolder)
                        {
                            drivers[i].AM = line.CellText.Trim();
                        }
                        if (driverLine + 19 == placeHolder)
                        {
                            drivers[i].PM = line.CellText.Trim();
                        }
                        if (driverLine + 20 == placeHolder)
                        {
                            drivers[i].onroadHours = line.CellText.Trim();
                        }
                        if (driverLine + 21 == placeHolder)
                        {
                            drivers[i].NDPPH = line.CellText.Trim();
                        }
                        if (driverLine + 22 == placeHolder)
                        {
                            drivers[i].helperHours = line.CellText.Trim();
                            i++;
                            placeHolder++;
                            driverLine = placeHolder;
                        }
                        else
                        {
                            placeHolder++;
                        }
                    }
                }

                if (!File.Exists("30daydriverData.json"))
                {
                    string jsonDrivers = JsonConvert.SerializeObject(drivers, Newtonsoft.Json.Formatting.Indented);
                    File.WriteAllText("30daydriverData.json", jsonDrivers);
                }
                else
                {
                    File.Delete("30daydriverData.json");
                    string jsonDrivers = JsonConvert.SerializeObject(drivers, Newtonsoft.Json.Formatting.Indented);
                    File.WriteAllText("30daydriverData.json", jsonDrivers);
                }

                MessageBox.Show($"Driver data exported successfully!", "Success", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.None);

            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}", "Error!", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        public void getdriverThirtyDay(string EmployeeID)
        {
            Boolean detected = false;
            int placeHolder = 0;
            int driverLine = 0;
            int i = 0;

            try
            {
                WebClient wClient = new WebClient();
                var htmlData = wClient.DownloadString($"https://pft.inside.ups.com/apps/DPSReports/SPMonthDtl.aspx?SLIC={sLIC.Text}&EMPID={EmployeeID}&BLDG={bLDG.Text}");
                HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
                hDoc.LoadHtml(htmlData);

                var tableQuery = from table in hDoc.DocumentNode.SelectNodes("//table").Cast<HtmlNode>()
                                 from row in table.SelectNodes("tr").Cast<HtmlNode>()
                                 from cell in row.SelectNodes("th|td").Cast<HtmlNode>()
                                 select new { Table = table.Id, CellText = cell.InnerText };

                List<thirtydaydriverDatawDate> drivers = new List<thirtydaydriverDatawDate>();

                foreach (var line in tableQuery)
                {

                    if (line.CellText.ToString() == "Rte")
                    {
                        detected = true;
                        driverLine = placeHolder + 18;
                    }
                    else
                    {
                        if (detected == false)
                        {
                            placeHolder++;
                        }
                    }
                    if (detected == true)
                    {
                        if (driverLine == placeHolder)
                        {
                            var driverA = new thirtydaydriverDatawDate() { driverName = line.CellText.Trim() };
                            drivers.Add(driverA);
                        }
                        if (driverLine + 1 == placeHolder)
                        {
                            drivers[i].employeeID = line.CellText.Trim();
                        }
                        if (driverLine + 2 == placeHolder)
                        {
                            drivers[i].supGroup = line.CellText.Trim();
                        }
                        if (driverLine + 3 == placeHolder)
                        {
                            drivers[i].paycode = line.CellText.Trim();
                        }
                        if (driverLine + 4 == placeHolder)
                        {
                            drivers[i].delDate = line.CellText.Trim();
                        }
                        if (driverLine + 5 == placeHolder)
                        {
                            drivers[i].route = line.CellText.Trim();
                        }
                        if (driverLine + 6 == placeHolder)
                        {
                            drivers[i].planDay = line.CellText.Trim();
                        }
                        if (driverLine + 7 == placeHolder)
                        {
                            drivers[i].OvUn = line.CellText.Trim();
                        }
                        if (driverLine + 8 == placeHolder)
                        {
                            drivers[i].paidDay = line.CellText.Trim();
                        }
                        if (driverLine + 9 == placeHolder)
                        {
                            drivers[i].SPORH = line.CellText.Trim();
                        }
                        if (driverLine + 10 == placeHolder)
                        {
                            drivers[i].miles = line.CellText.Trim();
                        }
                        if (driverLine + 11 == placeHolder)
                        {
                            drivers[i].SPM = line.CellText.Trim();
                        }
                        if (driverLine + 12 == placeHolder)
                        {
                            drivers[i].delStops = line.CellText.Trim();
                        }
                        if (driverLine + 13 == placeHolder)
                        {
                            drivers[i].puStops = line.CellText.Trim();
                        }
                        if (driverLine + 14 == placeHolder)
                        {
                            drivers[i].delPkgs = line.CellText.Trim();
                        }
                        if (driverLine + 15 == placeHolder)
                        {
                            drivers[i].puPkgs = line.CellText.Trim();
                        }
                        if (driverLine + 16 == placeHolder)
                        {
                            drivers[i].paidSApercent = line.CellText.Trim();
                        }
                        if (driverLine + 17 == placeHolder)
                        {
                            drivers[i].ccsendAgains = line.CellText.Trim();
                        }
                        if (driverLine + 18 == placeHolder)
                        {
                            drivers[i].AM = line.CellText.Trim();
                        }
                        if (driverLine + 19 == placeHolder)
                        {
                            drivers[i].PM = line.CellText.Trim();
                        }
                        if (driverLine + 20 == placeHolder)
                        {
                            drivers[i].onroadHours = line.CellText.Trim();
                        }
                        if (driverLine + 21 == placeHolder)
                        {
                            drivers[i].NDPPH = line.CellText.Trim();
                        }
                        if (driverLine + 22 == placeHolder)
                        {
                            drivers[i].helperHours = line.CellText.Trim();
                            i++;
                            placeHolder++;
                            driverLine = placeHolder;
                        }
                        else
                        {
                            placeHolder++;
                        }
                    }

                }

                if (!File.Exists("30dayDriverDetail.json"))
                {
                    string jsonDrivers = JsonConvert.SerializeObject(drivers, Newtonsoft.Json.Formatting.Indented);
                    File.WriteAllText("30dayDriverDetail.json", jsonDrivers);
                }
                else
                {
                    File.Delete("30dayDriverDetail.json");
                    string jsonDrivers = JsonConvert.SerializeObject(drivers, Newtonsoft.Json.Formatting.Indented);
                    File.WriteAllText("30dayDriverDetail.json", jsonDrivers);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}");
            }
        }
        public void driverData()
        {
            Boolean detected = false;
            int placeHolder = 0;
            int driverLine = 0;
            int i = 0;
            try
            {
                WebClient wClient = new WebClient();
                var data = wClient.DownloadString($"https://pft.inside.ups.com/apps/DPSReports/SPperfomanceTrendMonthAvg.aspx?BLDG={bLDG.Text}&SLIC={sLIC.Text}&GRID=False");
                HtmlAgilityPack.HtmlDocument hDoc = new HtmlAgilityPack.HtmlDocument();
                hDoc.LoadHtml(data);


                var tableQuery = from table in hDoc.DocumentNode.SelectNodes("//table").Cast<HtmlNode>()
                                 from row in table.SelectNodes("tr").Cast<HtmlNode>()
                                 from cell in row.SelectNodes("th|td").Cast<HtmlNode>()
                                 select new { Table = table.Id, CellText = cell.InnerText };

                List<drvrData> drivers = new List<drvrData>();

                foreach (var line in tableQuery)
                {

                    if (line.CellText.ToString() == "Rte")
                    {
                        detected = true;
                        driverLine = placeHolder + 17;
                    }
                    else
                    {
                        if (detected == false)
                        {
                            placeHolder++;
                        }
                    }

                    if (detected == true)
                    {
                        if (driverLine == placeHolder)
                        {
                            var driverA = new drvrData() { Route = line.CellText.Trim() };
                            drivers.Add(driverA);
                        }
                        if (driverLine + 1 == placeHolder)
                        {
                            drivers[i].Drivername = line.CellText.Trim();
                        }
                        if (driverLine + 2 == placeHolder)
                        {
                            drivers[i].EmployeeNumber = line.CellText.Trim();
                        }
                        if (driverLine + 3 == placeHolder)
                        {
                            drivers[i].DelStops = line.CellText.Trim();
                        }
                        if (driverLine + 4 == placeHolder)
                        {
                            drivers[i].DelPkgs = line.CellText.Trim();
                            //13
                        }
                        if (driverLine + 5 == placeHolder)
                        {
                            drivers[i].PUStops = line.CellText.Trim();
                        }
                        if (driverLine + 6 == placeHolder)
                        {
                            drivers[i].PUPkgs = line.CellText.Trim();
                        }
                        if (driverLine + 7 == placeHolder)
                        {
                            drivers[i].Miles = line.CellText.Trim();
                        }
                        if (driverLine + 8 == placeHolder)
                        {
                            drivers[i].SPM = line.CellText.Trim();
                        }
                        if (driverLine + 9 == placeHolder)
                        {
                            drivers[i].PlanDay = line.CellText.Trim();
                        }
                        if (driverLine + 10 == placeHolder)
                        {
                            drivers[i].PaidDay = line.CellText.Trim();
                        }
                        if (driverLine + 11 == placeHolder)
                        {
                            drivers[i].OvUn = line.CellText.Trim();
                        }
                        if (driverLine + 12 == placeHolder)
                        {
                            drivers[i].SPORH = line.CellText.Trim();
                        }
                        if (driverLine + 13 == placeHolder)
                        {
                            drivers[i].NDPPH = line.CellText.Trim();
                        }
                        if (driverLine + 14 == placeHolder)
                        {
                            drivers[i].AM = line.CellText.Trim();
                        }
                        if (driverLine + 15 == placeHolder)
                        {
                            drivers[i].PM = line.CellText.Trim();
                        }
                        if (driverLine + 16 == placeHolder)
                        {
                            drivers[i].NumDays = line.CellText.Trim();
                            i++;
                            placeHolder++;
                            driverLine = placeHolder;
                        }
                        else
                        {
                            placeHolder++;
                        }
                    }
                }

                if (!File.Exists("driverData.json"))
                {
                    string jsonDrivers = JsonConvert.SerializeObject(drivers, Newtonsoft.Json.Formatting.Indented);
                    File.WriteAllText("driverData.json", jsonDrivers);
                }
                else
                {
                    File.Delete("driverData.json");
                    string jsonDrivers = JsonConvert.SerializeObject(drivers, Newtonsoft.Json.Formatting.Indented);
                    File.WriteAllText("driverData.json", jsonDrivers);
                }

                MessageBox.Show($"Driver data exported successfully!", "Success", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.None);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERROR! \n {ex.Message}", "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void deleteAll()
        {
            planDay.Text = "";
            paidDay.Text = "";
            sPORH.Text = "";
            ovUn.Text = "";
            delPkgs.Text = "";
            delStops.Text = "";
            puPkgs.Text = "";
            puStops.Text = "";
            amTime.Text = "";
            pmTime.Text = "";
            employeeID.Text = "";
            NDPPH.Text = "";
            mILES.Text = "";
            SPM.Text = "";
            listBox1.Items.Clear();
        }

        public void deleteAll2()
        {
            planDay.Text = "";
            paidDay.Text = "";
            sPORH.Text = "";
            ovUn.Text = "";
            delPkgs.Text = "";
            delStops.Text = "";
            puPkgs.Text = "";
            puStops.Text = "";
            amTime.Text = "";
            pmTime.Text = "";
            employeeID.Text = "";
            NDPPH.Text = "";
            mILES.Text = "";
            SPM.Text = "";
        }

        public void clear30day()
        {
            delPkgs30.Text = "";
            delStops30.Text = "";
            puPkgs30.Text = "";
            puStops30.Text = "";
            ndpph30.Text = "";
            employeeID30.Text = "";
            daysonRoute30.Text = "";
            planDay30.Text = "";
            paidDay30.Text = "";
            ovUn30.Text = "";
            sporh30.Text = "";
            miles30.Text = "";
            spm30.Text = "";
            sendagainper30.Text = "";
            amTime30.Text = "";
            pmTime30.Text = "";
            onroadHours30.Text = "";
            helperHours30.Text = "";
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string data = File.ReadAllText("driverData.json");
            var driverData = JsonConvert.DeserializeObject<drvrData[]>(data);

            comboBox2.Items.Clear();
            foreach (var line in driverData)
            {
                if (line.Drivername == comboBox1.Text)
                {
                    comboBox2.Items.Add(line.Route);
                    comboBox2.SelectedIndex = 0;
                }
            }

            
            foreach (var line2 in driverData)
            {
                if (line2.Route == comboBox2.Text)
                {
                    if (line2.Drivername == comboBox1.Text)
                    {
                        deleteAll();
                        planDay.Text = line2.PlanDay;
                        paidDay.Text = line2.PaidDay;
                        ovUn.Text = line2.OvUn;
                        sPORH.Text = line2.SPORH;
                        amTime.Text = line2.AM;
                        pmTime.Text = line2.PM;
                        delStops.Text = line2.DelStops;
                        delPkgs.Text = line2.DelPkgs;
                        puPkgs.Text = line2.PUPkgs;
                        puStops.Text = line2.PUStops;
                        employeeID.Text = line2.EmployeeNumber;
                        NDPPH.Text = line2.NDPPH;
                        SPM.Text = line2.SPM;
                        mILES.Text = line2.Miles;
                        dateData();
                        break;
                    }
                    else
                    {
                        deleteAll();
                    }
                }
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            driverData();
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            string data = File.ReadAllText("driverData.json");
            var driverData2 = JsonConvert.DeserializeObject<drvrData[]>(data);
            foreach (var line in driverData2)
            {
                comboBox2.Items.Add(line.Route);
                comboBox1.Items.Add(line.Drivername);
            }
            object[] dItems = (from object o in comboBox2.Items select o).Distinct().ToArray();
            comboBox2.Items.Clear();
            comboBox2.Items.AddRange(dItems);
            object[] dItems2 = (from object o in comboBox1.Items select o).Distinct().ToArray();
            comboBox1.Items.Clear();
            comboBox1.Items.AddRange(dItems2);
            comboBox2.SelectedIndex = 0;
            comboBox1.SelectedIndex = 0;
            var DataTime = File.GetLastWriteTime("driverData.json");
            dteTime.Text = "";
            dteTime.Text = DataTime.ToString();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!File.Exists("driverData.json"))
            {
                MessageBox.Show("No driver data found! \n Initializing data pull....", "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
                driverData();
            }
            else
            {
                comboBox1.Items.Clear();
                comboBox2.Items.Clear();
                string data = File.ReadAllText("driverData.json");
                var driverData = JsonConvert.DeserializeObject<drvrData[]>(data);
                foreach (var line in driverData)
                {
                    comboBox2.Items.Add(line.Route);
                    comboBox1.Items.Add(line.Drivername);
                }
                object[] dItems = (from object o in comboBox2.Items select o).Distinct().ToArray();
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(dItems);
                object[] dItems2 = (from object o in comboBox1.Items select o).Distinct().ToArray();
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(dItems2);
                comboBox2.SelectedIndex= 0;
                comboBox1.SelectedIndex= 0;
                var DataTime = File.GetLastWriteTime("driverData.json");
                dteTime.Text = "";
                dteTime.Text = DataTime.ToString();
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string data = File.ReadAllText("driverData.json");
            var driverData = JsonConvert.DeserializeObject<drvrData[]>(data);
            
            foreach (var line in driverData)
            {
                if (line.Drivername == comboBox1.Text)
                {
                    if (line.Route == comboBox2.Text)
                    {
                        deleteAll();
                        planDay.Text = line.PlanDay;
                        paidDay.Text = line.PaidDay;
                        ovUn.Text = line.OvUn;
                        sPORH.Text = line.SPORH;
                        amTime.Text = line.AM;
                        pmTime.Text = line.PM;
                        delStops.Text = line.DelStops;
                        delPkgs.Text = line.DelPkgs;
                        puPkgs.Text = line.PUPkgs;
                        puStops.Text = line.PUStops;
                        employeeID.Text = line.EmployeeNumber;
                        NDPPH.Text = line.NDPPH;
                        SPM.Text = line.SPM;
                        mILES.Text = line.Miles;
                        dateData();
                        break;
                    }
                    else
                    {
                        deleteAll();
                    }
                }
            }            
        }

        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.WindowState= FormWindowState.Minimized;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (File.Exists("config.json"))
            {
                string jsonFile = File.ReadAllText("config.json");
                var data = JsonConvert.DeserializeObject<mainConfig>(jsonFile);
                bLDG.Text = data.Building;
                sLIC.Text = data.SLIC;
            }

            if (!File.Exists("30daydriverData.json"))
            {
                //nothing the file has been deleted or never generated.
            }
            else
            {
                driverthirty.Items.Clear();
                string driverthirtyData = File.ReadAllText("30daydriverData.json");
                var driverthirty2 = JsonConvert.DeserializeObject <thirtydaydriverData[]> (driverthirtyData);
                foreach (var line2 in driverthirty2)
                {
                    driverthirty.Items.Add(line2.driverName);
                }
                object[] dItems2 = (from object o in driverthirty.Items select o).Distinct().ToArray();
                driverthirty.Items.Clear();
                driverthirty.Items.AddRange(dItems2);
                driverthirty.SelectedIndex = 0;
                var DataTime2 = File.GetLastWriteTime("30daydriverData.json");
                thirdaydteTime.Text = "";
                thirdaydteTime.Text = DataTime2.ToString();
            }

            if (!File.Exists("driverData.json"))
            {
                //nothing... more than likely file does not exist or first run
            }
            else
            {
                comboBox1.Items.Clear();
                comboBox2.Items.Clear();
                string data = File.ReadAllText("driverData.json");
                var driverData = JsonConvert.DeserializeObject<drvrData[]>(data);
                foreach (var line in driverData)
                {
                    comboBox2.Items.Add(line.Route);
                    comboBox1.Items.Add(line.Drivername);
                }
                object[] dItems = (from object o in comboBox2.Items select o).Distinct().ToArray();
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(dItems);
                object[] dItems2 = (from object o in comboBox1.Items select o).Distinct().ToArray();
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(dItems2);
                comboBox2.SelectedIndex = 0;
                comboBox1.SelectedIndex = 0;
                var DataTime = File.GetLastWriteTime("driverData.json");
                dteTime.Text = "";
                dteTime.Text = DataTime.ToString();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            deleteAll2();
            string data = File.ReadAllText("driverdateData.json");
            var driverData = JsonConvert.DeserializeObject<dateDriver[]>(data);

            foreach (var date in driverData)
            {
                if (date.Date == listBox1.SelectedItem.ToString())
                {
                    delStops.Text = date.DelStops;
                    delPkgs.Text = date.DelPkgs;
                    puStops.Text = date.PUStops;
                    puPkgs.Text = date.PUPkgs;
                    mILES.Text = date.Miles;
                    SPM.Text = date.SPM;
                    NDPPH.Text = date.NDPPH;
                    planDay.Text = date.PlanDay;
                    paidDay.Text = date.PaidDay;
                    ovUn.Text = date.OvUn;
                    sPORH.Text = date.SPORH;
                    amTime.Text = date.AM;
                    pmTime.Text = date.PM;
                    employeeID.Text = date.EmployeeID;
                }
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (!File.Exists("driverdateData.json")){ }else{ File.Delete("driverdateData.json");}
            if (!File.Exists("30dayDriverDetail.json")) { } else { File.Delete("30dayDriverDetail.json"); }
            if (!File.Exists("config.json"))
            {
                mainConfig mConfig = new mainConfig();
                mConfig.Building = bLDG.Text.Trim();
                mConfig.SLIC = sLIC.Text.Trim();
                var data = JsonConvert.SerializeObject(mConfig, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText("config.json", data);
            }
            else
            {
                string jsonFile = File.ReadAllText("config.json");
                var data = JsonConvert.DeserializeObject<mainConfig>(jsonFile);
                data.Building = bLDG.Text.Trim();
                data.SLIC = sLIC.Text.Trim();
                var wdata = JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText("config.json", wdata);
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            thirtydayData();
            driverthirty.Items.Clear();
            string data = File.ReadAllText("30daydriverData.json");
            var driverData2 = JsonConvert.DeserializeObject<thirtydaydriverDatawDate[] >(data);
            foreach (var line in driverData2)
            {
                driverthirty.Items.Add(line.driverName);
            }
            object[] dItems = (from object o in driverthirty.Items select o).Distinct().ToArray();
            driverthirty.Items.Clear();
            driverthirty.Items.AddRange(dItems);
            driverthirty.SelectedIndex = 0;
            var DataTime = File.GetLastWriteTime("30daydriverData.json");
            //dteTime.Text = "";
            thirdaydteTime.Text = DataTime.ToString();
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void puPkgs_TextChanged(object sender, EventArgs e)
        {

        }

        private void driverthirty_SelectedIndexChanged(object sender, EventArgs e)
        {
            Boolean found = false;
            //https://pft.inside.ups.com/apps/DPSReports/SPMonthDtl.aspx?SLIC=1166&EMPID=3249536%20%20%20%20&BLDG=NYSPR
            try
            {
                listBox2.Items.Clear();
                routeThirty.Items.Clear();
                supThirty.Items.Clear();
                string data = File.ReadAllText("30daydriverData.json");
                var driverData = JsonConvert.DeserializeObject<thirtydaydriverData[]>(data);

                foreach (var line in driverData)
                {
                    if (line.driverName == driverthirty.Text)
                    {
                        getdriverThirtyDay(line.employeeID);
                        string drvrDetail = File.ReadAllText("30dayDriverDetail.json");
                        var driverDetail = JsonConvert.DeserializeObject<thirtydaydriverDatawDate[]>(drvrDetail);
                        foreach (var dtl in driverDetail)
                        {
                            listBox2.Items.Add(dtl.delDate);
                            routeThirty.Items.Add(dtl.route);
                            supThirty.Items.Add(dtl.supGroup);
                        }
                        found = true;
                    }
                    if (found == true)
                    {
                        break;
                    }
                }
                object[] dItems = (from object o in routeThirty.Items select o).Distinct().ToArray();
                routeThirty.Items.Clear();
                routeThirty.Items.AddRange(dItems);
                routeThirty.SelectedIndex = 0;
                object[] dItems2 = (from object o in supThirty.Items select o).Distinct().ToArray();
                supThirty.Items.Clear();
                supThirty.Items.AddRange(dItems2);
                supThirty.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}", "Error!", buttons: MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> list = new List<string>();
            string drvrDetail = File.ReadAllText("30dayDriverDetail.json");
            var driverDetail = JsonConvert.DeserializeObject<thirtydaydriverDatawDate[]>(drvrDetail);
            foreach (var dtl in driverDetail)
            {
                if (listBox2.SelectedItem.ToString() == dtl.delDate)
                {
                    clear30day();
                    list.Add($"Route: {dtl.route}");
                    list.Add($"Employee: {dtl.driverName}");
                    list.Add($"Employee ID: {dtl.employeeID}");
                    list.Add($"Supervisor Group: {dtl.supGroup}");
                    list.Add($"Plan Day: {dtl.planDay}");
                    list.Add($"Paid Day: {dtl.paidDay}");
                    list.Add($"Ov/Un: {dtl.OvUn}");
                    list.Add($"On Road Hours: {dtl.onroadHours}");
                    list.Add($"Delivery Stops: {dtl.delStops}");
                    list.Add($"Delivery Packages: {dtl.delPkgs}");
                    list.Add($"Pickup Stops: {dtl.puStops}");
                    list.Add($"Pickup Packages: {dtl.puPkgs}");
                    list.Add($"SPORH: {dtl.SPORH}");
                    list.Add($"Miles: {dtl.miles}");
                    list.Add($"AM Time: {dtl.AM}");
                    list.Add($"PM Time: {dtl.PM}");
                    var message = string.Join(Environment.NewLine, list);
                    MessageBox.Show(message, $"{dtl.driverName} - {listBox2.Text}", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Information);
                }
            }
        }

        private void routeThirty_SelectedIndexChanged(object sender, EventArgs e)
        {
            //supThirty.Items.Clear();
            List<string> list = new List<string>();
            string drvrInfo = File.ReadAllText("30daydriverData.json");
            var driver30Info = JsonConvert.DeserializeObject<thirtydaydriverData[]>(drvrInfo);
            foreach (var dtl in driver30Info)
            {
                if (dtl.driverName == driverthirty.Text)
                {
                    if (dtl.route == routeThirty.Text)
                    {
                        if (dtl.supGroup == supThirty.Text)
                        {
                            clear30day();
                            employeeID30.Text = dtl.employeeID.ToString();
                            daysonRoute30.Text = dtl.dayWork;
                            paidDay30.Text = dtl.paidDay.ToString();
                            planDay30.Text = dtl.planDay.ToString();
                            ovUn30.Text = dtl.OvUn;
                            sporh30.Text = dtl.SPORH;
                            spm30.Text = dtl.SPM;
                            miles30.Text = dtl.miles;
                            delPkgs30.Text = dtl.delPkgs;
                            delStops30.Text = dtl.delStops;
                            puPkgs30.Text = dtl.puPkgs;
                            puStops30.Text = dtl.puStops;
                            sendagainper30.Text = dtl.paidSApercent;
                            amTime30.Text = dtl.AM;
                            pmTime30.Text = dtl.PM;
                            onroadHours30.Text = dtl.onroadHours;
                            ndpph30.Text = dtl.NDPPH;
                            helperHours30.Text = dtl.helperHours;
                            /*
                            list.Add($"Driver: {dtl.driverName}");
                            list.Add($"Employee ID: {dtl.employeeID}");
                            list.Add($"Sup Group: {dtl.supGroup}");
                            list.Add($"PayCode: {dtl.paycode}");
                            list.Add($"Days on Route: {dtl.dayWork}");
                            list.Add($"Route: {dtl.route}");
                            list.Add($"Paid Day: {dtl.paidDay}");
                            list.Add($"Plan Day: {dtl.planDay}");
                            list.Add($"Ov/Un: {dtl.OvUn}");
                            list.Add($"SPORH: {dtl.SPORH}");
                            list.Add($"Miles: {dtl.miles}");
                            list.Add($"SPM: {dtl.SPM}");
                            list.Add($"Del Stops: {dtl.delStops}");
                            list.Add($"Del Pkgs: {dtl.delPkgs}");
                            list.Add($"PU Stops: {dtl.puStops}");
                            list.Add($"PU Pkgs: {dtl.puPkgs}");
                            list.Add($"Paid SA Percent: {dtl.paidSApercent}");
                            list.Add($"AM: {dtl.AM}");
                            list.Add($"PM: {dtl.PM}");
                            list.Add($"On Road Hours: {dtl.onroadHours}");
                            list.Add($"NDPPH: {dtl.NDPPH}");
                            list.Add($"Helper Hours: {dtl.helperHours}");
                            var message = string.Join(Environment.NewLine, list);
                            MessageBox.Show(message, "Driver Infomation by Day", MessageBoxButtons.OK, icon: MessageBoxIcon.Information);
                            */
                        }
                    }
                }
            }
        }

        private void supThirty_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> list = new List<string>();
            string drvrInfo = File.ReadAllText("30daydriverData.json");
            var driver30Info = JsonConvert.DeserializeObject<thirtydaydriverData[]>(drvrInfo);
            foreach (var dtl in driver30Info)
            {
                if (dtl.driverName == driverthirty.Text)
                {
                    if (dtl.route == routeThirty.Text)
                    {
                        if (dtl.supGroup == supThirty.Text)
                        {
                            /*
                            list.Add($"Driver: {dtl.driverName}");
                            list.Add($"Employee ID: {dtl.employeeID}");
                            list.Add($"Sup Group: {dtl.supGroup}");
                            list.Add($"PayCode: {dtl.paycode}");
                            list.Add($"Days on Route: {dtl.dayWork}");
                            list.Add($"Route: {dtl.route}");
                            list.Add($"Paid Day: {dtl.paidDay}");
                            list.Add($"Plan Day: {dtl.planDay}");
                            list.Add($"Ov/Un: {dtl.OvUn}");
                            list.Add($"SPORH: {dtl.SPORH}");
                            list.Add($"Miles: {dtl.miles}");
                            list.Add($"SPM: {dtl.SPM}");
                            list.Add($"Del Stops: {dtl.delStops}");
                            list.Add($"Del Pkgs: {dtl.delPkgs}");
                            list.Add($"PU Stops: {dtl.puStops}");
                            list.Add($"PU Pkgs: {dtl.puPkgs}");
                            list.Add($"Paid SA Percent: {dtl.paidSApercent}");
                            list.Add($"AM: {dtl.AM}");
                            list.Add($"PM: {dtl.PM}");
                            list.Add($"On Road Hours: {dtl.onroadHours}");
                            list.Add($"NDPPH: {dtl.NDPPH}");
                            list.Add($"Helper Hours: {dtl.helperHours}");
                            var message = string.Join(Environment.NewLine, list);
                            MessageBox.Show(message);
                            */
                            clear30day();
                            employeeID30.Text = dtl.employeeID.ToString();
                            daysonRoute30.Text = dtl.dayWork;
                            paidDay30.Text = dtl.paidDay.ToString();
                            planDay30.Text = dtl.planDay.ToString();
                            ovUn30.Text = dtl.OvUn;
                            sporh30.Text = dtl.SPORH;
                            spm30.Text = dtl.SPM;
                            miles30.Text = dtl.miles;
                            delPkgs30.Text = dtl.delPkgs;
                            delStops30.Text = dtl.delStops;
                            puPkgs30.Text = dtl.puPkgs;
                            puStops30.Text = dtl.puStops;
                            sendagainper30.Text = dtl.paidSApercent;
                            amTime30.Text = dtl.AM;
                            pmTime30.Text = dtl.PM;
                            onroadHours30.Text = dtl.onroadHours;
                            ndpph30.Text = dtl.NDPPH;
                            helperHours30.Text = dtl.helperHours;
                        }
                    }
                }
            }
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void orssPull_Click(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            WebClient wClient = new WebClient();
            string pull = wClient.DownloadString($"{textBox1.Text}");
            var jsonCon = JsonConvert.DeserializeObject<ORSS[]>(pull);

            foreach ( var item in jsonCon)
            {
                listBox3.Items.Add($"{item.RouteName} - {item.LastName}, {item.FirstName} - Sup Group: {item.PlannedTimecard.SupGroup}");
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> list = new List<string>();
            WebClient wClient = new WebClient();
            string pull = wClient.DownloadString($"{textBox1.Text}");
            var jsonCon = JsonConvert.DeserializeObject<ORSS[]>(pull);
            foreach (var item in jsonCon)
            {
                if (listBox3.Text == $"{item.LastName}, {item.FirstName}")
                {
                    list.Add($"Supervisor Group: {item.SupervisorGroup}");
                    list.Add($"Pay Code: {item.PayCode}");
                    list.Add($"Driver Type: {item.DriverTypeCode}");
                    list.Add($"Nickname: {item.Nickname}");
                    list.Add($"Employee ID: {item.ID}");
                    var message = string.Join(Environment.NewLine, list);
                    MessageBox.Show(message, $"{item.LastName}, {item.FirstName} Information", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Information);
                }
            }
        }
    }
}
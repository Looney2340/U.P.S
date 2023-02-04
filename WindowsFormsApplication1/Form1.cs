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

namespace WindowsFormsApplication1
{    public partial class Form1 : Form
    {
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
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
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string data = File.ReadAllText("driverData.json");
            var driverData = JsonConvert.DeserializeObject<drvrData[]>(data);
            foreach (var line in driverData)
            {
                if (line.Route == comboBox2.Text)
                {
                    if (line.Drivername == comboBox1.Text)
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
    }
}

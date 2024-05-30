using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml.Schema;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using Org.BouncyCastle.Utilities.Net;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace Getting_Company_Info
{
    public partial class Form1 : Form
    {
        List<string> alphabetlist = new List<string>();
        List<string> Companylist = new List<string>();
        List<string> datalist = new List<string>(); int j = 0;
        int alpha_a = 0; string STEP = "Step_1"; HtmlDocument theDoc = null; int t = 0;
        int company_a = 0; int skip = 0; int insert = 0; int duplicate = 0; MySqlDataReader myDataReader;

        MySqlConnection con = new MySqlConnection(@"Persist Security Info=False;User ID=root;pwd=;Initial Catalog=company_info;Data Source=localhost; charset=utf8;");

        public Form1()
        {
            InitializeComponent();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
            Environment.Exit(0);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Alphalink_Navigate();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/others");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/A");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/B");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/C");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/D");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/E");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/F");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/G");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/H");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/I");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/J");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/K");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/L");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/M");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/N");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/O");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/P");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/Q");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/R");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/S");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/T");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/U");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/V");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/W");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/X");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/Y");
            //alphabetlist.Add("https://www.moneycontrol.com/india/stockpricequote/Z");

            Alphatotallbl.Visible = true;
            Alphatotallbl.Text = "Total : " + alphabetlist.Count;

            webBrowser1.ScriptErrorsSuppressed = true;
        }
        public void Alphalink_Navigate()
        {
            if(alpha_a < alphabetlist.Count)
            {
                string link = alphabetlist[alpha_a].Trim();
                textBox1.Text = link;

                //string data = GetSource(link);

                webBrowser1.Navigate(link);
                timer2.Enabled = true;
                STEP = "Step_1";
                alpha_a++;

                alphacoplbl.Visible = true;
                alphacoplbl.Text = "Working : " + alpha_a + "/" + alphabetlist.Count;
                alphacoplbl.Refresh();
            }
            else
            {
                companylink_Navigate();
            }
        }
        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if(webBrowser1.ReadyState == WebBrowserReadyState.Complete)
            {
                theDoc = webBrowser1.Document;
                if(theDoc != null)
                {
                    switch (STEP)
                    {
                        case "Step_1":
                            timer2.Enabled = true;
                            break;

                        case "Step_2":
                            timer1.Enabled = true;                            
                            break;

                        case "Step_3":
                            break;
                    }
                }
                else
                {
                    Application.DoEvents();
                }                
            }
            else
            {
                Application.DoEvents();
            }
        } 
        public void CollectLinks()
        {
            try
            {
                HtmlElementCollection tblcoll = theDoc.GetElementsByTagName("TABLE");
                foreach (HtmlElement ele in tblcoll)
                {
                    if (ele.OuterHtml.Contains(">Company Name<"))
                    {
                        HtmlElementCollection tdcoll = theDoc.GetElementsByTagName("TD");
                        foreach (HtmlElement ele1 in tdcoll)
                        {
                            if(ele1.OuterHtml.Contains("<A") && !ele1.OuterHtml.Contains("></A>"))
                            {
                                HtmlElementCollection acoll = ele1.GetElementsByTagName("a");
                                string companylink = acoll[0].GetAttribute("href");
                                companylink = companylink.Replace("&amp;", "&");

                                Companylist.Add(companylink);

                                comlinklbl.Visible = true;
                                comlinklbl.Text = "Company : " + Companylist.Count;
                                comlinklbl.Refresh();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error In CollectLinks.", Application.ProductName,MessageBoxButtons.OK, MessageBoxIcon.Error);
                //Application.Exit();
                //Environment.Exit(0);
            }
        }
        public void companylink_Navigate()
        {
            if (company_a < Companylist.Count)
            {
                string link = Companylist[company_a].Trim();
                textBox1.Text = link; 
                webBrowser1.Navigate(link);
                STEP = "Step_2";
                timer1.Enabled = true;
                company_a++;
                //System.Threading.Thread.Sleep(3000); changes

                compcomplbl.Visible = true;
                compcomplbl.Text = "Working : " + company_a + "/" + Companylist.Count;
                compcomplbl.Refresh();
            }
            else
            {
                //InsertRecord();
            }
        }
        public void CollectData()
        {
            try
            {                
                string data = theDoc.Body.OuterHtml;
                if(!data.Contains("No Data For Registered Office."))
                {
                    if(data.Contains("Registered Office"))
                    {
                        string name = "";
                        string sector = "";

                        HtmlElement divcoll = theDoc.GetElementById("stockName");
                        HtmlElementCollection h1 = divcoll.GetElementsByTagName("H1");
                        HtmlElementCollection span = divcoll.GetElementsByTagName("SPAN");

                        if (h1[0].InnerText != null)
                        {
                            name = h1[0].InnerText.Trim();
                        }
                        if (span[0].InnerText != null)
                        {
                            sector = span[0].InnerText.Trim();
                            if (sector.Contains(":"))
                            {
                                sector = sector.Substring(sector.IndexOf(":") + 1).Trim();
                            }
                        }

                        HtmlElementCollection licoll = theDoc.GetElementsByTagName("LI");
                        foreach (HtmlElement ele in licoll)
                        {
                            if (ele.OuterHtml.Contains("Registered Office") && ele.OuterHtml.Contains("<p>"))
                            {
                                HtmlElementCollection pcoll = ele.GetElementsByTagName("P");
                                string address = "";
                                string city = "";
                                string state = "";
                                string pin = "";
                                string tel = "";
                                string fax = "";
                                string email = "";
                                string website = "";
                                if (pcoll[0].InnerText != null)
                                {
                                    address = pcoll[0].InnerText.Trim();
                                }
                                if (pcoll[01].InnerText != null)
                                {
                                    city = pcoll[01].InnerText.Trim();
                                }
                                if (pcoll[02].InnerText != null)
                                {
                                    state = pcoll[02].InnerText.Trim();
                                }
                                if (pcoll[03].InnerText != null)
                                {
                                    pin = pcoll[03].InnerText.Trim();
                                }
                                if (pcoll[04].InnerText != null)
                                {
                                    tel = pcoll[04].InnerText.Trim();
                                }
                                if (pcoll[05].InnerText != null)
                                {
                                    fax = pcoll[05].InnerText.Trim();
                                }
                                if (pcoll[06].InnerText != null)
                                {
                                    email = pcoll[06].InnerText.Trim();
                                }
                                if (pcoll[07].InnerText != null)
                                {
                                    website = pcoll[07].InnerText.Trim();
                                }

                                //string combineall = name + "☺" + sector + "☻" + address + "♥" + city + "♦" + state + "♣" + pin + "♠" + tel + "•" + fax + "◘" + email + "○" + website;
                                string combineall = name + "♥" + sector + "♥" + address + "♥" + city + "♥" + state + "♥" + pin + "♥" + tel + "♥" + fax + "♥" + email + "♥" + website;
                                combineall = combineall.Replace("&amp;", "&").Trim();
                                combineall = combineall.Replace(",,", ",").Trim();

                                InsertRecord(combineall);

                                datalist.Add(combineall);

                                collectlbl.Visible = true;
                                collectlbl.Text = "Collected : " + datalist.Count();
                                collectlbl.Refresh();
                                break;
                            }
                        }
                    }
                    else
                    {
                        skip++;
                        skiplbl.Visible = true;
                        skiplbl.Text = "Skipped : " + skip;
                        skiplbl.Refresh();
                    }                    
                }
                else
                {
                    skip++;
                    skiplbl.Visible = true;
                    skiplbl.Text = "Skipped : " + skip;
                    skiplbl.Refresh();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error In CollectData.", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                //Application.Exit();
                //Environment.Exit(0);
            }
        }
        public void InsertRecord(string combineall)
        {
            try
            {
                //for (int i = 0; i < datalist.Count; i++)
                //{
                if(combineall.Contains("'"))
                {
                    combineall = combineall.Replace("'", "''").Trim();
                }
                string[] arr = combineall.Split('♥');

                string name = arr[00].Trim();
                string sector = arr[01].Trim();
                string address = arr[02].Trim();
                string city = arr[03].Trim();
                string state = arr[04].Trim();
                string pin =arr[05].Trim();
                string tel = arr[06].Trim();
                string fax = arr[07].Trim();
                string email =arr[08].Trim();
                string website = arr[09].Trim();
                    
                if(address.EndsWith(","))
                {
                    address = address.Remove(address.LastIndexOf(",")).Trim();
                }
                    
                myDataReader = checkduplicate(email);
                if (myDataReader.HasRows == true)
                {
                    con.Close();
                    duplicate++;

                    duplbl.Visible = true;
                    duplbl.Text = "Duplicate : " + duplicate;
                    duplbl.Refresh();
                }
                else
                {
                    con.Close();
                    //MySqlConnection con = new MySqlConnection(@"Persist Security Info=False;User ID=root;pwd=;Initial Catalog=company_info;Data Source=localhost; charset=utf8;");
                    if (con.State != ConnectionState.Open)
                    {
                        con.Open();
                    }
                    string query = "INSERT INTO company_record (name, sector, address, city, state, pin, tel, fax, email, website) VALUES " +
                        "('" + name + "','" + sector + "','" + address + "','" + city + "','" + state + "','" + pin + "','" + tel + "','" + fax + "','" + email + "','" + website + "')";
                    MySqlCommand cmd = new MySqlCommand(query, con);
                    cmd.ExecuteNonQuery();
                    con.Close();

                    insert++;

                    insertlbl.Visible = true;
                    insertlbl.Text = "Inserted : " + insert;
                    insertlbl.Refresh();
                }
                //}
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error In Insert Record.", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                //Application.Exit();
                //Environment.Exit(0);
            }
        }
        public MySqlDataReader checkduplicate(string email)
        {
            
            try
            {
                if(con.State != ConnectionState.Open)
                {
                    con.Open();
                }
                string query = "SELECT email FROM company_record WHERE email = '" + email + "'";
                using (MySqlCommand cmd = new MySqlCommand())
                {
                    cmd.CommandText = query;
                    cmd.Connection = con;
                    myDataReader = cmd.ExecuteReader();
                }  
            }
            catch
            {

            }
            return myDataReader;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (t >= 25)
            {
                theDoc = webBrowser1.Document;
                t = 0;
                timer1.Enabled = false;
                if (theDoc != null)
                {
                    if (theDoc.Body != null)
                    {
                        if (theDoc.Body.InnerHtml.Contains("Registered Office"))
                        {
                            CollectData();
                            companylink_Navigate();
                        }
                        else
                        {
                            //timer1.Enabled = true;
                            if (j > 10)
                            {
                                company_a--;
                                companylink_Navigate();
                            }
                            else
                            {
                                j++;
                                timer1.Enabled = true;
                            }
                        }
                    }
                    else
                    {
                        timer1.Enabled = true;
                    }
                }
                else
                {
                    timer1.Enabled = true;
                }
            }
            else
            {
                t++;
                timer_lbl.Visible = true;
                timer_lbl.Text = (t * 2) + "% Loading...";
                timer_lbl.Refresh();
            }
        }

        private string GetSource(string url)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            System.Net.WebRequest WReq = null;
            System.Net.WebResponse WRes = null;
            System.IO.StreamReader SReader = null;
            WRes = null;
            string searchData = "";
            try
            {
                ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
                WReq = System.Net.WebRequest.Create(url);
                WReq.Timeout = 2 * 60 * 1000;
                WRes = WReq.GetResponse();
                SReader = new System.IO.StreamReader(WRes.GetResponseStream());
                searchData = SReader.ReadToEnd();
                WRes.Close();
                WReq = null;
                WRes = null;
            }
            catch
            {
                //GetSource(url);
            }
            return searchData;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (t >= 50)
            {
                theDoc = webBrowser1.Document;
                t = 0;
                timer1.Enabled = false;
                if (theDoc != null)
                {
                    if (theDoc.Body != null)
                    {
                        if (theDoc.Body.InnerHtml.Contains(">Company Name<"))
                        {
                            CollectLinks();
                            Alphalink_Navigate();
                        }
                        else
                        {
                            timer1.Enabled = true;
                        }
                    }
                    else
                    {
                        timer1.Enabled = true;
                    }
                }
                else
                {
                    timer1.Enabled = true;
                }
            }
            else
            {
                t++;
                timer_lbl.Visible = true;
                timer_lbl.Text = (t * 2) + "% Loading...";
                timer_lbl.Refresh();
            }
        }
    }
}

using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Management;
using System.Net.NetworkInformation;
using System.Net;
using Microsoft.Win32;
using System.IO;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;
using System.Drawing;

namespace BuildApp
{
    public partial class BuildApp : Form
    {
        //private string instanceName; 

        public BuildApp()
        {
            InitializeComponent();

            GetIPAddress();
            GetOs();
            GetRam();
            GetPcName();
            GetDns();
            GetInstalledSoftware();
            GetLocalUsers();
            GetgrpBoxInst();
            GetTotalHDDSize();
            GetVerList();
        }


        #region AddDns
        private void AddDns(object sender, EventArgs e)
        {
            ManagementClass mClass = new ManagementClass("Win32_NetworkAdapterConfiguration");
            ManagementObjectCollection mObjCol = mClass.GetInstances();

            bool errorOnTry = false;

            foreach (ManagementObject mObj in mObjCol)
            {
                if ((bool)mObj["IPEnabled"])
                {
                    ManagementBaseObject mboDns = mObj.GetMethodParameters("SetDNSServerSearchOrder");
                    if (mboDns != null)
                    {
                        //Assume X.X.X.X and X.X.X.X are the IPs.
                        string[] sIPs = { "172.26.49.130", "172.26.49.131" };

                        try
                        {
                            mboDns["DNSServerSearchOrder"] = sIPs;
                            mObj.InvokeMethod("SetDNSServerSearchOrder", mboDns, null);
                        }
                        catch
                        {
                            errorOnTry = true;
                        }
                        finally
                        {
                            if (errorOnTry)
                            {
                                labelReNameSatus.Text = (@"Failed");
                                chkbxdnsad.Checked = false;
                            }
                            else
                            {
                                labelReNameSatus.Text = (@"Address added");
                                chkbxdnsad.Checked = true;
                            }
                        }
                    }
                }
            }
        }
        #endregion

        #region GetPcName
        private void GetPcName()
        {
            //Get PC Name
            textBoxPCName.Text = System.Environment.MachineName;
            textBoxHostName.Text = (@"Hostname: " + System.Environment.MachineName);
            txtBoxSrv.Text = System.Environment.MachineName;
        }
        #endregion

        #region InstalledSoftware
        private void GetInstalledSoftware()
        {
            RegistryKey localMachine = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office");
            if (localMachine != null)
            {
                foreach (string key in localMachine.GetSubKeyNames())
                {
                    if (key == "14.0")
                        cB2010.Checked = true;
                    if (key == "15.0")
                        cB2010.Checked = true;
                    if (key == "16.0")
                        cB2010.Checked = true;
                }
            }
            RegistryKey sqllocalMachine = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\MICROSOFT");
            if (sqllocalMachine != null)
            {
                foreach (string sqlkey in sqllocalMachine.GetSubKeyNames())
                {
                    if (sqlkey == "Microsoft SQL Server")
                        cBSQL.Checked = true;
                }
            }

            RegistryTools regTools = new RegistryTools();
            cBALBCLI.Checked = regTools.IsApplictionInstalled("ALB Client");
            cBALBSER.Checked = regTools.IsApplictionInstalled("ALB Server");
            cBMLC.Checked = regTools.IsApplictionInstalled("MLC");

            string saproot = "C:\\Program Files (x86)\\SAP BusinessObjects";
            if (Directory.Exists(saproot))
                cBSAP.Checked = true;
            else cBSAP.Checked = false;

            string lfmroot = "C:\\Program Files (x86)\\Laserform";
            if (Directory.Exists(lfmroot))
                cBLFM.Checked = true;
            else cBLFM.Checked = false;
        }
        #endregion

        #region DNS
        private void GetDns()
        {
            //DNS
            StringBuilder stringBuilder = new StringBuilder(string.Empty);

            NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();
            foreach (NetworkInterface ni in nics)
            {
                if (ni.OperationalStatus == OperationalStatus.Up)
                {
                    IPAddressCollection ips = ni.GetIPProperties().DnsAddresses;
                    foreach (System.Net.IPAddress dns in ips)
                    {
                        if (dns.ToString().Length < 16)
                        {
                            stringBuilder.AppendLine(dns.ToString());
                        }
                    }
                }
            }
            rtxtbxdns01.Text = stringBuilder.ToString();
            rtxtbxdns01.SelectionAlignment = HorizontalAlignment.Center;
        }
        #endregion

        #region RAM
        private void GetRam()
        {
            //RAM
            ManagementObjectSearcher Search = new ManagementObjectSearcher("Select * From Win32_ComputerSystem");

            string ramSize = "";

            foreach (ManagementObject Mobject in Search.Get())
            {
                double ramBytes = (Convert.ToDouble(Mobject["TotalPhysicalMemory"]));
                double inGigab = ramBytes / 1073741824;

                if (inGigab > 1)
                    ramSize = Convert.ToString(Math.Round(inGigab, 2)) + "GB";
                else
                    ramSize = Convert.ToString(Math.Round(inGigab * 1000, 2)) + "MB";

            }
            tBRAM.Text = ramSize;
        }
        #endregion

        #region OSVersion
        private void GetOs()
        {
            //OS
            RegistryKey keyAppRoot = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows NT\CurrentVersion");
            tBOS.Text = (string)keyAppRoot.GetValue("ProductName");

            //Processor
            ManagementObjectSearcher mosProcessor = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");
            string Procname = null;

            foreach (ManagementObject moProcessor in mosProcessor.Get())
            {
                if (moProcessor["name"] != null)
                {
                    Procname = moProcessor["name"].ToString();
                    tBProc.Text = Procname;
                }
            }
        }
        #endregion

        #region IpAddresses
        private void GetIPAddress()
        {
            //IP address 
            IPHostEntry host;
            string localIp = "IP Address";
            host = Dns.GetHostEntry(Dns.GetHostName());
            foreach (IPAddress ip in host.AddressList)
            {
                if (ip.AddressFamily.ToString() == "InterNetwork")
                {
                    localIp = ip.ToString();

                }
                txtbxIPadd.Text = localIp;
            }
        }
        #endregion

        #region CloseApp
        private void CloseApp(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion

        #region ShutDown
        private void ShutDown(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("shutdown.exe", "-r -t 0");
        }
        #endregion

        #region DomainJoin
        private void DomainJoin(object sender, EventArgs e)
        {
            string domain = "mlctest";
            string password = "Passw0rd1";
            string username = "Install";

            object[] methodArgs = { domain, password, username, null, 3 };

            ManagementObject computerSystem = new ManagementObject("Win32_ComputerSystem.Name='" + Environment.MachineName + "'");

            object oresult = computerSystem.InvokeMethod("JoinDomainOrWorkgroup", methodArgs);

            int result = (int)Convert.ToInt32(oresult);

            //list of errors
            string strConsoleOutput = " ";
            switch (result)
            {
                case 0:
                    strConsoleOutput = "Joined Successfully!";
                    System.Diagnostics.Process.Start("shutdown.exe", "-r -t 0");
                    break;
                case 5:
                    strConsoleOutput = "Access is denied";
                    break;
                case 87:
                    strConsoleOutput = "The parameter is incorrect";
                    break;
                case 110:
                    strConsoleOutput = "The system cannot open the specified object";
                    break;
                case 1323:
                    strConsoleOutput = "Unable to update the password";
                    break;
                case 1326:
                    strConsoleOutput = "Logon failure: unknown username or bad password";
                    break;
                case 1355:
                    strConsoleOutput = "The specified domain either does not exist or could not be contacted";
                    break;
                case 2224:
                    strConsoleOutput = "The account already exists";
                    break;
                case 2691:
                    strConsoleOutput = "The machine is already joined to the domain";
                    break;
                case 2692:
                    strConsoleOutput = "The machine is not currently joined to a domain";
                    break;
                default:
                    strConsoleOutput = "Error code is " + result;
                    break;
            }
            //Console.WriteLine(strConsoleOutput);
            lbldom.Text = (strConsoleOutput);
            chkbxdom.Checked = true;
            return;
        }
        #endregion

        #region PCName
        private void btnGetPC_Click(object sender, EventArgs e)
        {
            textBoxPCName.Text = System.Environment.MachineName;
        }
        #endregion

        #region DNSAddress
        private void IPAddress(object sender, EventArgs e)
        {
            StringBuilder StringBuilder1 = new StringBuilder(string.Empty);

            NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();
            foreach (NetworkInterface ni in nics)
            {
                if (ni.OperationalStatus == OperationalStatus.Up)
                {
                    IPAddressCollection ips = ni.GetIPProperties().DnsAddresses;
                    foreach (System.Net.IPAddress ip in ips)
                    {
                        if (ip.ToString().Length < 16)
                        {
                            StringBuilder1.AppendLine(ip.ToString());
                        }
                    }
                }
            }
            rtxtbxdns01.Text = StringBuilder1.ToString();
        }
        #endregion

        #region DisableIP6
        private void btndisip6_Click(object sender, EventArgs e)
        {
            Microsoft.Win32.RegistryKey myKey;
            myKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Services\\tcpip6\\Parameters", true);
            if (myKey != null)
            {
                myKey.SetValue("DisabledComponents", "FFFFFFFF");
                myKey.Close();
                lblip6.Text = "Reboot required";
            }
        }
        #endregion

        #region PCName
        private void btnSetPCName_Click(object sender, EventArgs e)

        {
            if (textBoxPCName.Text != "")
            {
                string compPath = "Win32_ComputerSystem.Name='" + System.Environment.MachineName + "'";

                using (ManagementObject dir = new ManagementObject(new ManagementPath(compPath)))
                {
                    ManagementBaseObject inputArgs = dir.GetMethodParameters("Rename");
                    inputArgs["Name"] = textBoxPCName.Text;
                    try
                    {
                        ManagementBaseObject outParams = dir.InvokeMethod("Rename", inputArgs, null);
                    }
                    catch
                    {
                        lbldnsprog.Text = ("Failed");
                    }
                    finally
                    {
                        lbldnsprog.Text = ("PC name changed, reboot if not joining to domain");
                        chkbxpcrename.Checked = true;
                    }
                }
            }
        }
        #endregion

        #region LocalUsers
        private void GetLocalUsers()
        {
            ManagementObjectSearcher usersSearcher = new ManagementObjectSearcher(@"SELECT * FROM Win32_UserAccount WHERE Domain = '" + System.Environment.MachineName + "'");
            ManagementObjectCollection users = usersSearcher.Get();

            var localUsers = users.Cast<ManagementObject>().Where(
                u => (bool)u["LocalAccount"] == true &&
                     (bool)u["Disabled"] == false &&
                     (bool)u["Lockout"] == false &&
                     int.Parse(u["SIDType"].ToString()) == 1 &&
                     u["Name"].ToString() != "HomeGroupUser$");

            foreach (ManagementObject user in users)
            {
                lstboxLU.Items.Add("Local Account: " + user["Name"].ToString());
            }
        }
        #endregion

        #region Office2016Fix
        private void Office2016Fix(object sender, EventArgs e)
        {
            string SourcePath = @"\\lglsan\Software\IT\Outlook 2016 Autodiscover fix\Auto";
            string DestinationPath = @"C:\Auto";
            string[] directories = System.IO.Directory.GetDirectories(SourcePath, " *.*", SearchOption.AllDirectories);
            if (Directory.Exists(DestinationPath))
            {
                Parallel.ForEach(directories, dirPath => { Directory.CreateDirectory(dirPath.Replace(SourcePath, DestinationPath)); });
                string[] files = System.IO.Directory.GetFiles(SourcePath, "*.*", SearchOption.AllDirectories);
                Parallel.ForEach(files, newPath => { File.Copy(newPath, newPath.Replace(SourcePath, DestinationPath)); });
            }
            else
            {
                Directory.CreateDirectory(DestinationPath);
                Parallel.ForEach(directories, dirPath => { Directory.CreateDirectory(dirPath.Replace(SourcePath, DestinationPath)); });
                string[] files = System.IO.Directory.GetFiles(SourcePath, "*.*", SearchOption.AllDirectories);
                Parallel.ForEach(files, newPath => { File.Copy(newPath, newPath.Replace(SourcePath, DestinationPath)); });
            }
            Process regeditProcess = Process.Start("regedit.exe", "/s c:\\auto\\Autodiscover.reg");
            regeditProcess.WaitForExit();

            lblBoxO16Auto.Text = ("Fix Installed");
        }
        #endregion

        #region SQLInstanceName
        public void GetgrpBoxInst()
        {
            RegistryView registryView = Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32;
            using (RegistryKey hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, registryView))
            {
                RegistryKey instanceKey32 = hklm.OpenSubKey(@"SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL", false);
                RegistryKey instanceKey64 = hklm.OpenSubKey(@"SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\Instance Names\SQL", false);
                if (instanceKey32 != null)
                {
                    foreach (var instanceName in instanceKey32.GetValueNames())
                    {
                        lstBoxInst.Items.Add(Environment.MachineName + @"\" + instanceName);
                    }
                }
                else if (instanceKey64 != null)
                {
                    foreach (var instanceName in instanceKey64.GetValueNames())
                    {
                        lstBoxInst.Items.Add(Environment.MachineName + @"\" + instanceName);
                    }
                }
            }
        }
        #endregion

        #region HDDSize
        private void GetTotalHDDSize()
        {
            DriveInfo[] allDrives = DriveInfo.GetDrives();

                foreach (DriveInfo d in allDrives)
            {
                lstBxDisks.Items.Add("Drive letter:   " + d.Name);
                if (d.IsReady == true)
                {
                    lstBxDisks.Items.Add("Total Size:    " + d.TotalSize / (1024 * 1024 * 1024) + "gb");
                    lstBxDisks.Items.Add("Free Space: " + d.TotalFreeSpace / (1024 * 1024 * 1024) + "gb");
                }
                else
                {
                    lstBxDisks.Items.Add("CD Drive");
                }
            }
        }
        #endregion

        #region SQLVersionList
        private void GetVerList()
        {
            RegistryKey localMachine = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Microsoft SQL Server\");
            if (localMachine != null)
            {

                foreach (string key in localMachine.GetSubKeyNames())
                    
                {
                    if (key == "130")
                        chkBox16.Checked = true;
                    if (key == "120")
                        chkBox14.Checked = true;
                    if (key == "110")
                        chkBox12.Checked = true;
                    if (key == "100")
                        chkBox08.Checked = true;
                    if (key == "90")
                        chkBox05.Checked = true;
                    if (key != null)
                        noSQLChkBx.Checked = false;
                }
            }
        }
        #endregion

        private void btnSQLGo_Click(object sender, EventArgs e)
        {
            RegistryView registryView = Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32;
            using (RegistryKey hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, registryView))

            {
                RegistryKey instanceKey32 = hklm.OpenSubKey(@"SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL", false);
                RegistryKey instanceKey64 = hklm.OpenSubKey(@"SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\Instance Names\SQL", false);
                

                DataTable databases = null;

                var connectionString = string.Format("Server="+ txtBoxSrv.Text +";User ID="+ txtBxUsrName.Text +";Password="+ txtBxPass.Text +"; Integrated Security=false");
                
                using (var sqlConnection = new SqlConnection(connectionString))
                {

                    if (instanceKey32 != null || instanceKey64 != null)
                    {
                        sqlConnection.Open();
                        databases = sqlConnection.GetSchema("Databases");
                        sqlConnection.Close();

                        if (databases != null)
                        {
                            foreach (DataRow row in databases.Rows)
                            {
                                lstBoxDB.Items.Add(row["database_name"]);
                            }
                        }
                    }
                    else
                    {
                        lstBoxDB.Items.Add("No database installed.");
                    }
                }
            }
        }

        private void label1_Click_2(object sender, EventArgs e)
        {

        }
        private void cB2010_CheckedChanged(object sender, EventArgs e)
        {

        }
        private void lstboxLU_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void cBALBSER_CheckedChanged(object sender, EventArgs e)
        {

        }
        private void tabPage1_Click(object sender, EventArgs e)
        {

        }
        private void rtxtbxdns01_TextChanged(object sender, EventArgs e)
        {
            rtxtbxdns01.SelectAll();
            rtxtbxdns01.SelectionAlignment = HorizontalAlignment.Center;
        }

        private void gBNetwork_Enter(object sender, EventArgs e)
        {

        }
        private void chkbxpcrename_CheckedChanged(object sender, EventArgs e)
        {

        }
        private void txtbxIPadd_TextChanged(object sender, EventArgs e)
        {
        }
        private void label1_Click(object sender, EventArgs e)
        {
        }
        private void textBoxPCName_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtBxPass_TextChanged(object sender, EventArgs e)
        {
            txtBxPass.UseSystemPasswordChar = true;

        }

        private void txtBoxSrv_TextChanged(object sender, EventArgs e)
        {
            txtBoxSrv.SelectAll();
            txtBoxSrv.TextAlign = HorizontalAlignment.Center;
        }

        private void textBoxHostName_TextChanged(object sender, EventArgs e)
        {

        }

        private void lstBxVMs_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnVM_Click(object sender, EventArgs e)
        {
            VMs vmsObj = new VMs();
            foreach (string vmItem in vmsObj.LocalVMs())
                lstBxVMs.Items.Add(vmItem);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            bool pingable = false;
            Ping pinger = new Ping();

            if (string.IsNullOrWhiteSpace(txtBoxSrvStat.Text))
            {
                MessageBox.Show("Server name missing", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
                this.timer1.Start();
                PingReply reply = pinger.Send(txtBoxSrvStat.Text);
                pingable = reply.Status == IPStatus.Success;
                pingable = reply.Status == IPStatus.DestinationUnreachable;
                pingable = reply.Status == IPStatus.TimedOut;

                if (reply.Status == IPStatus.Success)
                {
                    chkBoxOnline.Checked = true;
                    picBox.Enabled = true;
                    picBox.BackColor = Color.Green;
                }
                else if (reply.Status == IPStatus.DestinationUnreachable)
                {
                    picBox.Enabled = false;
                    chkBoxOffline.Checked = true;
                    picBoxOff.Enabled = true;
                    picBoxOff.BackColor = Color.Red;
                }
                else if (reply.Status == IPStatus.TimedOut)
                {
                    picBox.Enabled = false;
                    chkBoxOffline.Checked = true;
                    picBoxOff.Enabled = true;
                    picBoxOff.BackColor = Color.Red;
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtBoxSrvStat.Text = "";
            chkBoxOffline.Checked = false;
            chkBoxOnline.Checked = false;
            progressBar1.Value = 0;
            timer1.Enabled = false;
            picBox.Enabled = false;
            picBoxOff.Enabled = false;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.progressBar1.Increment(2);
        }
    }
}



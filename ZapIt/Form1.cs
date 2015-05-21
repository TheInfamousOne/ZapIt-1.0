using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;
using Microsoft.Win32;
using System.Diagnostics;
using System.Net;
using System.Net.NetworkInformation;



namespace ZapIt
{
    public partial class Form1 : Form
    {
        int switchexpression = 3;
        string osArch = "";
        OpenFileDialog ofd = new OpenFileDialog();
        string uniqueName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;


        public Form1()
        {
            InitializeComponent();
            textBox1.Select();
        }
        
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                textBox1.Font = new Font(textBox1.Font, FontStyle.Italic);
                textBox1.ForeColor = Color.LightGray;
            }
        }

        private void machineInfo()
        {
            string TERMID = textBox1.Text;
            RegistryKey regKey = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, TERMID).OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Department");

            var AssetTag = regKey.GetValue("AssetTag");
            var BIOSVendor = regKey.GetValue("BIOSVendor");
            var BIOSVersion = regKey.GetValue("BIOSVersion");
            var BootDiskVersion = regKey.GetValue("BootDiskVersion");
            var BuildDate = regKey.GetValue("BuildDate");
            var Building = regKey.GetValue("Building");
            var BuildingFloor = regKey.GetValue("BuildingFloor");
            var BuildingSection = regKey.GetValue("BuildingSection");
            var BuildTime = regKey.GetValue("BuildTime");
            var BuiltBy = regKey.GetValue("BuiltBy");
            var Model = regKey.GetValue("Model");
            var DepartmentName = regKey.GetValue("DepartmentName");
            var Encryption = regKey.GetValue("Encryption");
            var Mode = regKey.GetValue("Mode");
            var MaintLastBegin = regKey.GetValue("MaintLastBegin");
            var MaintLastFinish = regKey.GetValue("MaintLastFinish");
            var MaintLastSuccess = regKey.GetValue("MaintLastSuccess");


            listBox1.Items.Add("Mode: " + Mode);
            listBox1.Items.Add("AssetTag: " + AssetTag);
            listBox1.Items.Add("Location: " + Building + " " + BuildingSection + " FLOOR " + BuildingFloor);
            listBox1.Items.Add("Manufactuer: " + BIOSVendor);
            listBox1.Items.Add("Model: " + Model);
            listBox1.Items.Add("Bios Version: " + BIOSVersion);
            listBox1.Items.Add("BootDisk Version: " + BootDiskVersion);
            listBox1.Items.Add("Built By: " + BuiltBy);
            listBox1.Items.Add("Build Date: " + BuildDate);
            listBox1.Items.Add("Build Time: " + BuildTime);
            listBox1.Items.Add("Department: " + DepartmentName);
            listBox1.Items.Add("Encryption: " + Encryption);
            listBox1.Items.Add("Last Time Maintenance Began: " + MaintLastBegin);
            listBox1.Items.Add("Last Time Maintenance Finished " + MaintLastFinish);
            listBox1.Items.Add("Last Successful Maintenance: " + MaintLastSuccess);
            enableButtons();
        }

        private void WindowsInstaller()
        {
            string TERMID = textBox1.Text;
            string FileName = @"\\" + TERMID + @"\c$\WsMgmt\Logs\msizap.log";
            if (File.Exists(FileName))
            {
                File.Delete(FileName);
            }

            if (string.IsNullOrEmpty(textBox1.Text))
            {
                label1.Text = "Invalid TermId";
            }
            else
            {                
                object[] theNewProcessToRun = { @"C:\Windows\System32\cmd.exe /c C:\WSMGMT\Bin\WindowsInstaller.exe >> C:\WSMGMT\Logs\msizap.log" };
                //object[] theProcessToRun = { @"C:\WSMGMT\Bin\msizap.exe TW! {269ED7D8-AF8A-4208-9479-4862F492A0BA} > C:\WSMGMT\BIN\AA.log" };
                ManagementClass theClass = new ManagementClass(@"\\" + TERMID + @"\root\cimv2:Win32_Process");
                theClass.InvokeMethod("Create", theNewProcessToRun);
                isFileLocked();

            }

        }

        private void msiX()
        {
            string TERMID = textBox1.Text;
            string GUID = textBox2.Text;
            object[] theNewProcessToRun = { @"C:\Windows\System32\msiexec.exe /X" + GUID + @" /qb-! /Lime C:\WsMgmt\Logs\ZapItUninstall.log" };
            ManagementClass theClass = new ManagementClass(@"\\" + TERMID + @"\root\cimv2:Win32_Process");
            theClass.InvokeMethod("Create", theNewProcessToRun);
            readUninstallLog();


        }

        private void ZapIt()
        {
            string TERMID = textBox1.Text;
            string GUID = textBox2.Text;            
            object[] theNewProcessToRun = { @"C:\Windows\System32\cmd.exe /c C:\WSMGMT\Bin\msizap.exe TW! " + GUID + @">> C:\WsMgmt\Logs\ZapIt.log" };
            ManagementClass theClass = new ManagementClass(@"\\" + TERMID + @"\root\cimv2:Win32_Process");
            theClass.InvokeMethod("Create", theNewProcessToRun);
            readZapIt();
        }

        private void readZapIt()
        {
            string TERMID = textBox1.Text;
            string fileinuse = "inUse";            
            while (fileinuse != null)
            {
                this.Cursor = Cursors.WaitCursor;
                try
                {
                    string file = @"\\" + TERMID + @"\c$\WsMgmt\Logs\ZapIt.log";
                    using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                    {
                        using (StreamReader sr = new StreamReader(fs))
                        {
                            while (!sr.EndOfStream)
                            {
                                string result = sr.ReadToEnd();
                                textBox3.Text = result;
                                // Console.WriteLine(sr.ReadLine());
                                fileinuse = null;
                            }
                        }
                    }
                }

                catch (IOException)
                {
                    textBox3.Text = "File currently in use.";
                    fileinuse = "Streaming data MsiZap data.";
                }
            }
            this.Cursor = Cursors.Default;
            delLogFiles();
            disableInstallZapButtons();
        }
  
        private void readUninstallLog()
        {
            
            string TERMID = textBox1.Text;
            string fileinuse = "inUse";            
            while (fileinuse != null)
            {
                this.Cursor = Cursors.WaitCursor;
                try
                {
                    string file = @"\\" + TERMID + @"\c$\WsMgmt\Logs\ZapItUninstall.log";
                    using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                    {
                        using (StreamReader sr = new StreamReader(fs))
                        {
                            while (!sr.EndOfStream)
                            {
                                string result = sr.ReadToEnd();
                                textBox3.Text = result;
                                fileinuse = null;
                            }

                        }
                    }

                }

                catch (IOException)
                {
                    textBox3.Text = "File currently in use.";
                    fileinuse = "Streaming data MsiZap data.";
                }

            }
            this.Cursor = Cursors.Default;
            delLogFiles();
            disableInstallZapButtons();

        }

        private void delLogFiles()
        {
            string TERMID = textBox1.Text;
            DateTime time = DateTime.Now;


            if (File.Exists(@"\\" + TERMID + @"\c$\WsMgmt\Logs\MsiZap.log"))
            {
                File.Delete(@"\\" + TERMID + @"\c$\WsMgmt\Logs\MsiZap.log");
            }

            if (File.Exists(@"\\" + TERMID + @"\c$\WsMgmt\Logs\ZapItUninstall.log"))
            {
                string uninstallContents = File.ReadAllText(@"\\" + TERMID + @"\c$\WsMgmt\Logs\ZapItUninstall.log");
                string[] stampSession = { "**************************************************************", "UNINSTALLED BY ZAPIT UTILITY", "Uninstalled by " + uniqueName, "DATE:" + time.ToString(), "**************************************************************" };
                File.Delete(@"\\" + TERMID + @"\c$\WsMgmt\Logs\ZapItUninstall.log");
                File.AppendAllLines(@"\\" + TERMID + @"\c$\WsMgmt\Logs\ZapItMaster.log", stampSession);
                File.AppendAllText(@"\\" + TERMID + @"\c$\WsMgmt\Logs\ZapItMaster.log", uninstallContents);


            }

            if (File.Exists(@"\\" + TERMID + @"\c$\WsMgmt\Logs\ZapIt.log"))
            {
                string zapContents = File.ReadAllText(@"\\" + TERMID + @"\c$\WsMgmt\Logs\ZapIt.log");
                string[] stampSession = { "**************************************************************", "REGISTRY CLEANED BY ZAPIT UTILITY", "Zapped by " + uniqueName, "DATE:" + time.ToString(), "**************************************************************" };
                File.Delete(@"\\" + TERMID + @"\c$\WsMgmt\Logs\ZapIt.log");              
                File.AppendAllLines(@"\\" + TERMID + @"\c$\WsMgmt\Logs\ZapItMaster.log", stampSession);
                File.AppendAllText(@"\\" + TERMID + @"\c$\WsMgmt\Logs\ZapItMaster.log", zapContents);

            }
            listBox1.Items.Clear();


        }

        private void hideListBox()
        {
            listBox1.Visible = false;
            textBox3.Visible = true;
        }

        private void isFileLocked()
        {
            string TERMID = textBox1.Text;
            bool inuse = true;
            this.Cursor = Cursors.WaitCursor;
            while (inuse == true)
            {
                try
                {
                    using (File.Open(@"\\" + TERMID + @"\c$\WsMgmt\Logs\msizap.log", FileMode.Open))
                    {                              
                        inuse = false;                        
                        break;
                    }
                }
                catch (IOException)
                {
                    
                }                
            }

          readFile();
        }

        private void readFile()
        {            
            string TERMID = textBox1.Text;
            string FileName = @"\\" + TERMID + @"\c$\WsMgmt\Logs\msizap.log";
            char[] chars = { '{', '}' };
            string characters = new string(chars);          
                                 
                   if (File.Exists(FileName))
                        {
                          using (StreamReader sr = new StreamReader(FileName))
                            {                                
                                while (!sr.EndOfStream)
                                {
                                    
                                    string line = sr.ReadToEnd();
                                    string[] readText = File.ReadAllLines(FileName);
                                    foreach (string FileText in readText)
                                    {                                        
                                        foreach (char c in characters)
                                        {
                                            if (FileText.Contains(c)) continue;
                                            listBox1.Sorted = true;
                                            listBox1.Items.Add(FileText);                                            
                                            break;
                                        }
                                    }                                    
                                 }
                              }                       
                                           
                          }
                   this.Cursor = Cursors.Default;                   
        }            
                
        private void disableInstallZapButtons()
        {
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
        }

        public void PingTest()
        {
            //clear listbox1 
            listBox1.Items.Clear();
            string TERMID = textBox1.Text;
            //set Ping time out
            int timeout = 120;
            //create ping object
            Ping ping = new Ping();
            //create ping buffer, datastream and byte
            PingOptions options = new PingOptions();
            string data = "searching for computer";
            byte[] buffer = Encoding.ASCII.GetBytes(data);

            try
            {
                PingReply reply = ping.Send(TERMID, timeout, buffer, options);
                if (reply.Status == IPStatus.Success)
                {
                    DirectoryExist();
                    label1.Text = "CONNECTED: " + TERMID + "  -  " + osArch;
                    switch (switchexpression)
                    {
                        //do nothing
                        case 0:
                            break;
                        //case1 will get Department Keys and display TEMRID info
                        case 1:
                            copyRequiredFiles();
                            machineInfo();
                            break;
                        //case2 will use WMI and call Win32_Products and get Software Installed
                        // case 2:
                        //   wmiGetSoftware();
                        //   break;
                    }
                }
                else
                {
                    label1.Text = TERMID + " is offline at the moment.";

                }

            }

            catch (PingException)
            {
                label1.Text = TERMID + " is offline, powered down or does not exist.";
            }
        }

        public void DirectoryExist()
        {
            string TERMID = textBox1.Text;
            listBox1.Items.Clear();
            if (Directory.Exists("\\\\" + TERMID + "\\c$\\Program Files (x86)"))
            {
                osArch = "Windows 7 64 bit";

            }
            else
            {
                osArch = "Windows 7 32 bit";
            }


        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {

            
            textBox1.ForeColor = Color.Black;
            textBox1.Font = new Font(textBox1.Font, FontStyle.Regular);

            if (e.KeyCode == Keys.Enter)
            {
                if (textBox1.Text.Length < 8)
                {
                    listBox1.Items.Clear();
                   
                    label1.Text = "Invalid TEMRID";

                }
                else
                {
                    pictureBox1.Visible = false;
                    listBox1.Items.Clear();                    
                    textBox2.Clear();
                    textBox3.Visible = false;
                    listBox1.Visible = true;
                    groupBox3.Text = "MSI Product Code";
                    switchexpression = 1;
                    PingTest();

                }
            }
        }

        private void enableButtons()
        {
            button1.Enabled = true;
            button4.Enabled = true;
        }

        private void enableUninstallZapButtons()
        {
            button1.Enabled = false;
            button2.Enabled = true;
            button3.Enabled = true;
        }

        private void copyRequiredFiles()
        {
            string TERMID = textBox1.Text;
            //if (File.Exists(@"C:\WSMGMT\Bin\streams.exe"))
            //{}
            //else
            //{ File.Copy("streams.exe", @"C:\WsMgmt\BIN\streams.exe"); }

            if (File.Exists("\\\\" + TERMID + "\\C$\\WSMGMT\\BIN\\MsiZap.exe"))
            { }
            else
            { File.Copy("MSIZAP.EXE", "\\\\" + TERMID + "\\c$\\WsMgmt\\BIN\\MsiZap.exe"); }

            if (File.Exists(@"C:\WsMgmt\BIN\psexec.exe"))
            { }
            else
            { File.Copy("psexec.exe", @"C:\WsMgmt\BIN\psexec.exe"); }

            if (File.Exists("\\\\" + TERMID + "\\C$\\WSMGMT\\BIN\\WindowsInstaller.exe"))
            { }
            else
            { File.Copy("WindowsInstaller.exe", "\\\\" + TERMID + "\\c$\\WSMGMT\\BIN\\WindowsInstaller.exe"); }

            if (File.Exists(@"\\" + TERMID + @"\C$\WsMgmt\BIN\Microsoft.Deployment.WindowsInstaller.dll"))
            { }
            else { File.Copy("Microsoft.Deployment.WindowsInstaller.dll", @"\\" + TERMID + @"\C$\WsMgmt\BIN\Microsoft.Deployment.WindowsInstaller.dll"); }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string TERMID = textBox1.Text;      
            if (File.Exists(@"\\" + TERMID + @"\C$\WsMgmt\Logs\msizap.log"))
            {
                File.Delete(@"\\" + TERMID + @"\C$\WsMgmt\Logs\msizap.log");
            }
            listBox1.Items.Clear();                        
            textBox2.Clear();
            textBox3.Clear();
            switchexpression = 0;
            groupBox3.Text = "MSI Product Code";
            button4.Enabled = true;
            WindowsInstaller();
        }

        public void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (switchexpression == 1) { }
            else
            {
                string softwareItem = listBox1.SelectedItem.ToString();
                textBox2.Text = "";
                string TERMID = textBox1.Text;

                using (StreamReader r = new StreamReader(@"\\" + TERMID + @"\c$\WsMgmt\Logs\msizap.log"))
                {
                    string line, version = "";
                    bool nameFound = false;
                    while ((line = r.ReadLine()) != null)
                    {
                        if (nameFound)
                        {
                            version = line;
                            textBox2.Text = version;
                            enableUninstallZapButtons();
                            break;
                        }

                        if (line.IndexOf(softwareItem) != -1)
                        {
                            nameFound = true;
                        }
                    }

                    if (version != "")
                    {
                        // version variable contains the product version
                    }
                    else
                    {
                        // not found
                    }
                }
            }

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string softwareItem = listBox1.SelectedItem.ToString();
            if (MessageBox.Show("Are you sure you would like to uninstall " + softwareItem + "?", "Warning", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                hideListBox();
                msiX();
            }
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            string TERMID = textBox1.Text;
            if (TERMID == "TERMID")
            {
                textBox1.Clear();
            }
            textBox2.Clear();
            textBox3.Clear();
            groupBox3.Text = "MSI Product Code";
            textBox3.Visible = false;
            listBox1.Visible = true;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string softwareItem = listBox1.SelectedItem.ToString();            
            if (MessageBox.Show("This will remove all registry entries for " + softwareItem + " but will not remove installation directory and files. Are you sure you would like to Zap " + softwareItem + "?", "Warning", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {                
                hideListBox();                
                ZapIt();
            }

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {

            if (textBox1.Text.Length == 8)
            {
                delLogFiles();
            }
            else { }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                label1.Text = "Invalid TermID";
            }
            else
            {
                disableInstallZapButtons();
                string TERMID = textBox1.Text;
                listBox1.Items.Clear();
                ofd.InitialDirectory = @"\\umhssmsps\Software Repository\Production";
                ofd.Filter = "Pick Executable|*.exe";


                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    textBox2.Text = ofd.SafeFileName;
                    string packagePath = ofd.FileName;
                    groupBox3.Text = "Installing Pacakge";

                    DialogResult Result = MessageBox.Show("Are you sure you would like to install" + ofd.FileName + " ?", "Pacakge Install", MessageBoxButtons.YesNo);
                    if (Result == DialogResult.Yes)
                    {
                        this.Cursor = Cursors.WaitCursor;
                        ProcessStartInfo startInfo;
                        startInfo = new ProcessStartInfo();
                        startInfo.FileName = "psexec.exe";
                        startInfo.WorkingDirectory = @"C:\WSMGMT\BIN\";
                        startInfo.Arguments = @"-h \\" + TERMID + @" -accepteula -u " + uniqueName + " cmd /c " + (" \"" + packagePath + "\"");
                        Process.Start(startInfo);
                    }
                    else
                    {
                        MessageBox.Show("Installation Cancelled.");
                    }
                    this.Cursor = Cursors.Default;
                    //delLogFiles();
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }
              
        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            pictureBox1.Visible = false;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

     
      


    }
}

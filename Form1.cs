using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UsbaPrinter
{
    public partial class Form1 : Form
    {
        static String path = "";
        static String serviceName = "PrintingService";
        private ServiceProcessInstaller process;
        private ServiceInstaller serviceInstaller1;

        public Form1()
        {
            InitializeComponent();
            path = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\service" + "\\PrintingService.exe";
        }

        String folderDirectory = "C:\\Users\\Public\\Documents" + "\\printer";

        private void btnSave_Click(object sender, EventArgs e)
        {
            string selected = this.A4sizePrinter.GetItemText(this.A4sizePrinter.SelectedItem);
            string selected2 = this.TharmalPrinter.GetItemText(this.TharmalPrinter.SelectedItem);


            string[] printer = { selected, selected2 };
            string output = string.Join("\n", printer);
            try
            {

                String FileName = folderDirectory + "\\printerdata.txt";
                if (!Directory.Exists(folderDirectory))
                {
                    Directory.CreateDirectory(folderDirectory);
                    saveInfo(FileName, output);
                }
                else
                {
                    saveInfo(FileName, output);
                }

                String urlFileName = folderDirectory + "\\url.txt";
                if (!Directory.Exists(folderDirectory))
                {
                    Directory.CreateDirectory(folderDirectory);
                    saveInfo(urlFileName, urlText.Text);
                }
                else
                {
                    saveInfo(urlFileName, urlText.Text);
                }


                InstallService(serviceName, Assembly.LoadFrom(path));

            }
            catch (Exception ex)
            { 
                Console.WriteLine(ex.Message);
            }
        }
        public static void InstallService(string serviceName, Assembly assembly)
        {


            using (AssemblyInstaller installer = GetInstaller(assembly))
            {
                //GetServiceAccount(installer)
                IDictionary state = new Hashtable();
                try
                {
                    installer.Install(state);
                    installer.Commit(state);
                    MessageBox.Show("Service started Successfully", "Printer Data ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch
                {
                    try
                    {
                        installer.Rollback(state);
                    }
                    catch (Exception e)
                    {

                    }
                    throw;
                }
            }

        }

        public static void startService()
        {

            if (IsServiceInstalled(serviceName))
            {
                ServiceController sc = new ServiceController(serviceName);
                if (sc.Status == ServiceControllerStatus.Stopped)
                {
                    sc.Start();
                }
                return;
            }
        }

        public static bool IsServiceInstalled(string serviceName)
        {
            using (ServiceController controller = new ServiceController(serviceName))
            {
                try
                {
                    ServiceControllerStatus status = controller.Status;
                }
                catch
                {
                    return false;
                }

                return true;
            }
        }

        private static AssemblyInstaller GetInstaller(Assembly assembly)
        {
            AssemblyInstaller installer = new AssemblyInstaller(assembly, null);
            installer.UseNewContext = true;

            return installer;
        }

        public void saveInfo(String path, String output) {
            try
            {
                using (StreamWriter sw = File.CreateText(path)) { sw.WriteLine(output); }
                

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            PrinterList();
           
        }

        private void PrinterList()
        {
             String FileName = folderDirectory + "/printerdata.txt";
            String printer1="", printer2="";

            if (File.Exists(FileName))
            {
                string[] lines = File.ReadAllLines(FileName);

                printer1 = lines[0].Trim();
                printer2 = lines[1].Trim();
                btnSave.Enabled = true;
            }
            else {
                btnSave.Enabled = false;
            }

            foreach (string sPrinters in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                A4sizePrinter.Items.Add(sPrinters);
                TharmalPrinter.Items.Add(sPrinters);


            }


            A4sizePrinter.SelectedIndex = A4sizePrinter.FindStringExact(printer1);
            TharmalPrinter.SelectedIndex = A4sizePrinter.FindStringExact(printer2);


        }

        private void btnCancel_Click(object sender, EventArgs e)
        {


            removeService();


            Application.Exit();
        }
        public async Task callUninstallerAsync() {
            await Task.Run(() => removeService());
        }

        private void PrinterSelect_Load(object sender, EventArgs e)
        {
            PrinterList();

            

        }

       

        private void A4sizePrinter_SelectedValueChanged_1(object sender, EventArgs e)
        {
            if (A4sizePrinter.Items == null)
            {
                btnSave.Enabled = false;
            }
            else {
                btnSave.Enabled = true;
            }
        }

        private void TharmalPrinter_SelectedValueChanged(object sender, EventArgs e)
        {
            if (TharmalPrinter.Items == null)
            {
                btnSave.Enabled = false;
            }
            else {
                btnSave.Enabled = true;
            }
            
        }

        private void label4_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }


        public void removeService()
        {
            try
            {
                ServiceInstaller ServiceInstallerObj = new ServiceInstaller();
                InstallContext Context = new InstallContext("<<log file path>>", null);
                ServiceInstallerObj.Context = Context;
                ServiceInstallerObj.ServiceName = serviceName;
                ServiceInstallerObj.Uninstall(null);
            }
            catch (Exception e) { 
            }

          

        }
    }
}

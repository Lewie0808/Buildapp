using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Management;
using System.Windows.Forms;

namespace BuildApp
{
    public class VMs
    {
        public List<string> LocalVMs()
        {
            RegistryView registryView = Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32;
            using (RegistryKey hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, registryView))

            {
                RegistryKey hostServer = hklm.OpenSubKey(@"SOFTWARE\Microsoft\Virtual Machine\Guest\Parameters", false);

                List<string> vmList = new List<string>();

                ManagementScope manScope = new ManagementScope(@"\\" + System.Environment.MachineName + @"\root\virtualization\v2");
                ObjectQuery queryObj = new ObjectQuery("SELECT * FROM Msvm_ComputerSystem");

                ManagementObjectSearcher vmSearcher = new ManagementObjectSearcher(manScope, queryObj);
                ManagementObjectCollection vmCollection = vmSearcher.Get();

                if (hostServer != null)
                {
                    vmList.Add("This is a virtual machine");
                    //MessageBox.Show("vm");
                }
                else
                {
                    //MessageBox.Show("real");
                    foreach (ManagementObject vm in vmCollection)
                    {
                        if (vm["ElementName"].ToString() != System.Environment.MachineName)
                        {
                            string vmOutput = vm["ElementName"] + " - " + vm["Status"];
                            vmList.Add("VM Name: " + vmOutput);
                        }
                    }
                }
                return vmList;
            }
        }
    }
}


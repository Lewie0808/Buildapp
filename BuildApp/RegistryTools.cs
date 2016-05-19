using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace BuildApp
{
    public class RegistryTools
    {
        private bool ExistsInSubKey(RegistryKey p_root, string p_subKeyName, string p_attributeName, string p_name)
        {
            RegistryKey subkey;
            string displayName;

            using (RegistryKey key = p_root.OpenSubKey(p_subKeyName))
            {
                if (key != null)
                {
                    foreach (string kn in key.GetSubKeyNames())
                    {
                        using (subkey = key.OpenSubKey(kn))
                        {
                            displayName = subkey.GetValue(p_attributeName) as string;
                            if (p_name.Equals(displayName, StringComparison.OrdinalIgnoreCase) == true)
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            return false;
        }

        public bool DeleteRegKeyTree(string keyPath)
        {
            try
            {
                Registry.LocalMachine.DeleteSubKeyTree(keyPath, true);
            }
            catch
            {
                return false;
            }
            return true;
        }

        public bool IsRootKeyPresent(string keyName)
        {
            RegistryKey rkSubKey = Registry.LocalMachine.OpenSubKey(keyName, false);
            if (rkSubKey == null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool IsApplictionInstalled(string applicationName)
        {
            string keyName;

            // search in: CurrentUser
            keyName = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            if (ExistsInSubKey(Registry.CurrentUser, keyName, "DisplayName", applicationName) == true)
            {
                return true;
            }

            // search in: LocalMachine_32
            keyName = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            if (ExistsInSubKey(Registry.LocalMachine, keyName, "DisplayName", applicationName) == true)
            {
                return true;
            }

            // search in: LocalMachine_64
            keyName = @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall";
            if (ExistsInSubKey(Registry.LocalMachine, keyName, "DisplayName", applicationName) == true)
            {
                return true;
            }

            return false;
        }

        public string LoadRegKey(string keyPath, string keyName)
        {
            string returnString = "";

            try
            {
                RegistryKey myKey = Registry.LocalMachine.OpenSubKey(keyPath, true);
                Object regKey = myKey.GetValue(keyName);
                returnString = regKey.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            return returnString;
        }

        public bool SaveRegKey(string keyPath, string keyName, string newValue)
        {
            bool status = true;

            try
            {
                RegistryKey myKey = Registry.LocalMachine.OpenSubKey(keyPath, true);
                myKey.SetValue(keyName, newValue, RegistryValueKind.String);
            }
            catch (Exception ex)
            {
                status = false;
                Console.WriteLine(ex.ToString());
            }

            return status;
        }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Net.NetworkInformation;
using System.Net;

namespace FacturacionElectronica
{
    public class Utils
    {

        // Lee variables del app.config
        public static string ReadSetting(string key, string defaultValue)
        {
            try
            {
                object setting = System.Configuration.ConfigurationSettings.AppSettings[key];
                if (setting != null)
                {
                    if (setting.ToString() == "")
                    {
                        setting = null;
                    }
                }
                return (setting == null) ? defaultValue : (string)setting;
            }
            catch
            {
                return defaultValue;
            }
        }

        // Obtiene la Mac de la pc
        public static ArrayList GetMacAddress()
        {
            ArrayList _array = new ArrayList();
            try
            {
                string macAddresses = "";
                string ip = "";

                foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
                {
                    //if (nic.NetworkInterfaceType != NetworkInterfaceType.Ethernet) continue;
                    if (nic.OperationalStatus == OperationalStatus.Up)
                    {
                        macAddresses += nic.GetPhysicalAddress().ToString();
                        ip += nic.GetIPProperties().UnicastAddresses[0].Address.ToString();

                        break;
                    }
                }

                if (ip.Length > 15)
                {
                    ip = "";
                    IPHostEntry host;
                    host = Dns.GetHostEntry(Dns.GetHostName());
                    foreach (IPAddress ipe in host.AddressList)
                    {
                        if (ipe.AddressFamily.ToString() == "InterNetwork")
                        {
                            ip = ipe.ToString();
                        }
                    }
                }

                _array.Add(macAddresses);
                _array.Add(ip);
                return _array;
            }
            catch (Exception ex)
            { }
            return _array;
        }
    }
}

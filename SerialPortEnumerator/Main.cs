using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;
using Wox.Plugin;

namespace SerialPortEnumerator
{
    public class Main : IPlugin
    {
        public void Init(PluginInitContext context) { }

        public List<Result> Query(Query query)
        {
            /*List<Result> res = SerialPort.GetPortNames().Select(serial => new Result()
            {
                Title = $"{serial}",
                SubTitle = "",
                IcoPath = "icon.png" //relative path to your plugin directory
            }).ToList();*/

            List<Result> res = new List<Result>();

            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(
                    "root\\CIMV2",
                    "SELECT * FROM Win32_PnPEntity WHERE ClassGuid=\"{4d36e978-e325-11ce-bfc1-08002be10318}\""
                );

                foreach (ManagementObject queryObj in searcher.Get())
                {
                    res.Add(new Result { Title = $"{queryObj["Caption"]}", SubTitle = "", IcoPath = "icon.png" });
                }
            }
            catch (ManagementException e)
            {
                
            }

            if (res.Count == 0)
                res.Add(new Result {Title = "No serial port",SubTitle = "", IcoPath = "icon.png" });
            return res;
        }
    }
}

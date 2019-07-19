using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;

namespace LogViewer_V2
{
    class logClass
    {
        public String logDate { get; set; }
        public String message { get; set; }
        public String level { get; set; }
        public String logType { get; set;}
        public String timeStamp { get; set; }
        public String fingerprint { get; set; }
        public String windowsIdentity { get; set; }
        public String machineName { get; set; }
        public String processName { get; set; }
        public String processVersion { get; set; }
        public String jobId { get; set; }
        public String robotName { get; set; }
        public String machineId { get; set; }
        public String fileName { get; set; }
        public String totalExecutionTimeInSeconds { get; set; }
        public String totalExecutionTime { get; set; }

        public List<string> GetAllPropertyValues()
        {
            List<string> values = new List<string>();
            foreach (var pi in typeof(logClass).GetProperties())
            {
                if (pi.GetValue(this,null) == null)
                {
                    values.Add("");
                }
                else
                {
                    values.Add(pi.GetValue(this, null).ToString());
                }
                
            }

            return values;
        }

    }
}

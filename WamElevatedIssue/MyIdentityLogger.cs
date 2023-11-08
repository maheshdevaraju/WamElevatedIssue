using Microsoft.IdentityModel.Abstractions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WamElevatedIssue
{
    class MyIdentityLogger : IIdentityLogger
    {
        private static readonly log4net.ILog log =
            log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public EventLogLevel MinLogLevel { get; }

        public MyIdentityLogger()
        {
            //Retrieve the log level from an environment variable
            var msalEnvLogLevel = Environment.GetEnvironmentVariable("MSAL_LOG_LEVEL");

            if (Enum.TryParse(msalEnvLogLevel, out EventLogLevel msalLogLevel))
            {
                MinLogLevel = msalLogLevel;
            }
            else
            {
                //Recommended default log level
                MinLogLevel = EventLogLevel.Verbose;
            }
        }

        public bool IsEnabled(EventLogLevel eventLogLevel)
        {
            return eventLogLevel <= MinLogLevel;
        }

        public void Log(LogEntry entry)
        {
            //Log Message here:
            log.Info(entry.Message);
            //Console.WriteLine(entry.Message);
        }
    }
}

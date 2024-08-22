using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EY.US.RecordAddin
{
    internal class Constants
    {
        public static readonly string mappingFile = "AddinConfig.xml";

        public static readonly string ClientRec = "Client Records";

        public static readonly string InternalFirmRec = "Internal Firm Records";

        public static readonly string logName = "ProcessLog";

        public static readonly string clientID = "Client Metadata ID";

        public static readonly string engID = "Eng Metadata ID";

        public static readonly string ClientName = "Client Name";

        public static readonly string engName = "Eng Name";

        public static readonly string popupMsg = "Please verify the following:\n\n";

        public static readonly string SaveCancel = "save operation cancelled";

        public static readonly string NoMatch = "No matching field found";

        public static readonly string MissingInfo = "Client or Engagement information cannot be blank for this record";

        public static void writelogInLocal(string message)
        {
            try
            {
                string path = "C:\\temp\\RecordAddinLogs.log";
                File.AppendAllText(path, DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "     " + message + Environment.NewLine);
            }
            catch (Exception)
            {
            }
        }
    }
}
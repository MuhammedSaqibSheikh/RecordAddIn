using EY.US.RecordAddin;
using TRIM.SDK;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestClient
{
    class Program
    {
        static void Main(string[] args)
        {
            Database db = new Database();
            db.WorkgroupServerName = "INMACHANDRAN03"; db.Id = "UD"; db.Connect();


            Record rec = new Record(db, new TrimURI(1326));
            RecordAddin objTest = new RecordAddin();
            objTest.Initialise(db);
            
            objTest.PreSave(rec);

            db.Dispose();
        }
    }
}

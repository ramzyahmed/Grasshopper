using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyGrasshopperAssembly1
{
    public class HardCodedData
    {

        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);


        // The source project (SAP_To_RVT) builds one folder per target framework; each .gha loads the matching one.
#if NETFRAMEWORK
        private const string SourceTargetFramework = "net48";
#else
        private const string SourceTargetFramework = "net8.0-windows";
#endif

        public static string Bim_directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Build", "BimIntegrationForGH", SourceTargetFramework);

    }
}

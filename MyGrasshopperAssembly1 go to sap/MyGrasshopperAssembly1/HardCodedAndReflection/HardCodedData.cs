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


        public static string Bim_directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Build", "BimIntegrationForGH");

    }
}

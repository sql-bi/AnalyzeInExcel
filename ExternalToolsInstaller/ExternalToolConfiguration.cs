using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExternalToolsInstaller
{
    public class ExternalToolConfiguration
    {
        public string version { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public string path { get; set; }
        public string arguments { get; set; }
        public string iconData { get; set; }
    }

    public class PathExternalToolConfiguration
    {
        public PathExternalToolConfiguration( string filename )
        {
            Filename = filename;
            string path = Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86);
            FullPath = System.IO.Path.Combine(path, @"Microsoft Shared\Power BI Desktop\External Tools\" + Filename);
        }

        public string Filename { get; }

        public string FullPath { get; }

    }
}

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Configuration.Internal;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace StockCSV.CordnersConfiguration
{
    public static class ConfigurationManager
    {
        static ConfigurationManager()
        {
            var configFile = new ExeConfigurationFileMap
            {
                ExeConfigFilename =
                    Path.Combine(Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath),
                        "Cordner.config")
            };
            General = System.Configuration.ConfigurationManager.OpenMappedExeConfiguration(configFile,
                ConfigurationUserLevel.None);
            Cordner = (CordnerConfigurationSection) General.GetSection("CordnersConfig");
        }

        public static Configuration General { get; }

        public static CordnerConfigurationSection Cordner { get; }
        
    }
}

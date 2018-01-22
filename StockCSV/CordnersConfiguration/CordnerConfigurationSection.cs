using System.Configuration;

namespace StockCSV.CordnersConfiguration
{
    public class CordnerConfigurationSection : ConfigurationSection
    {
        private static CordnerConfigurationSection cordnerConfig = (CordnerConfigurationSection)System.Configuration.ConfigurationManager.GetSection("CordnerConfigurationSection/cordnerConfig");

        private const string OutputPathProviderName = "Location";

        public static CordnerConfigurationSection CordnerConfig => cordnerConfig;

        [ConfigurationProperty(OutputPathProviderName)]
        public OutputConfigurationElement OutputLocation
        {
            get => (OutputConfigurationElement) this[OutputPathProviderName];
            set => this[OutputPathProviderName] = value;
        }
    }
}

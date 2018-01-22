using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StockCSV.CordnersConfiguration
{
    public class OutputConfigurationElement : ConfigurationElement
    {
        private const string LocationConfigurationPropertyName = "Location";

        [ConfigurationProperty(LocationConfigurationPropertyName, IsRequired = false)]
        public string Location
        {
            get => (string)this[LocationConfigurationPropertyName];
            set => this[LocationConfigurationPropertyName] = value;
        }
    }
}

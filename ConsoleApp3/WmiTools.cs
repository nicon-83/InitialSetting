using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;

namespace MVA.ConsoleApp
{
    class WmiTools
    {
        public static void PrintPropertiesOfWmiClass(string namespaceName, string wmiClassName)
        {
            ManagementPath managementPath = new ManagementPath();
            managementPath.Path = namespaceName;
            ManagementScope managementScope = new ManagementScope(managementPath);
            ObjectQuery objectQuery = new ObjectQuery("SELECT * FROM " + wmiClassName);
            ManagementObjectSearcher objectSearcher = new ManagementObjectSearcher(managementScope, objectQuery);
            ManagementObjectCollection objectCollection = objectSearcher.Get();
            foreach (ManagementObject managementObject in objectCollection)
            {
                PropertyDataCollection props = managementObject.Properties;
                foreach (PropertyData prop in props)
                {
                    Console.WriteLine("{0}:{1}", prop.Name, prop.Value);
                    //Console.WriteLine("Property type: {0}", prop.Type);
                    //Console.WriteLine("Property value: {0}", prop.Value);
                }
            }
        }

        public static void PrintPropertiesOfManagementObject(ManagementObject @object)
        {
            PropertyDataCollection data = @object.Properties;
            foreach (PropertyData item in data)
            {
                Console.WriteLine("{0}:\t{1}", item.Name, item.Value);
            }
        }
    }
}

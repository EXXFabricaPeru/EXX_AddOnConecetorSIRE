using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE
{
    public class SAPEntityNameAttribute : Attribute
    {
        public string Uri { get; }

        public SAPEntityNameAttribute(string uri)
        {
            Uri = uri;
        }
    }

    [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
    public class SAPKeyAttribute : Attribute
    {
    }
}

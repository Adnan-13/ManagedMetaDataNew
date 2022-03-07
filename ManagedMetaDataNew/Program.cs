using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ManagedMetaDataNew
{
    class Program
    {
        static void Main(string[] args)
        {
            Api api = new Api("https://genweb2bd.sharepoint.com/sites/Classic102", "Open");

            api.Test();
        }
    }
}

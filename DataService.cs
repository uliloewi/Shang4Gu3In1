using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shang4Gu3In1
{
    public class DataService
    {
        public static HttpClient Client = new HttpClient()
        {
            Timeout = Timeout.InfiniteTimeSpan
        };
    }
}

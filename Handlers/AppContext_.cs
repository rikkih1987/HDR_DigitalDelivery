using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HDR_EMEA;

namespace HDR_EMEA.Handlers
{
    public static class AppContext_
    {
        public static ExternalEvent ExternalEventModelHealth { get; set; }
        public static void RaiseModelHealthEvent()
        {
            ExternalEventModelHealth?.Raise();
        }
        
    }
}

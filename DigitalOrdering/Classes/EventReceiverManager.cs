using Microsoft.SharePoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DigitalOrdering.Classes
{
    public class EventReceiverManager : SPEventReceiverBase, IDisposable
    {
        public EventReceiverManager(bool disableImmediately)
        {
            EventFiringEnabled = !disableImmediately;
        }

        public void StopEventReceiver()
        {
            EventFiringEnabled = false;
        }
        public void StartEventReceiver()
        {
            EventFiringEnabled = true;
        }

        public void Dispose()
        {
            EventFiringEnabled = true;
        }
    }
}

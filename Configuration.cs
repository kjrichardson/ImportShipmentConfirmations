using System;
using System.Collections.Generic;
using System.Text;

namespace ImportShipmentConfirmations
{
    public class Configuration
    {

        public string ApiUser { get; set; }
        public string ApiPassword { get; set; }
        public string BaseURL { get; set; }
        public string GhostScriptFolder { get; set; }
        public ShipmentConfiguration Shipment { get; set; }

    }

    public class ShipmentConfiguration
    {
        public string Url { get; set; }
        public string InputFolder { get; set; }
        public string OutputFolder { get; set; }
        public string ProblemFolder { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace MicMon.Properties
{

    public record SelectedDevice(int DeviceNumber, string Id, string FriendlyName);

    public record class SelectedDevicesDocument()
    {
        [JsonInclude]
        public List<SelectedDevice> SelectedDevices = new List<SelectedDevice>();
    }
}

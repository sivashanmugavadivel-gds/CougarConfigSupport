using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CougarConfigSupport
{
    /////////////////////////////////////////////////////OPCUA Type//////////////////////////////////////////////////////
    public class OPCDictionaryDataType
    {
        public string name { get; set; }
        public string version { get; set; }
        public string comment { get; set; }
        public string instanceNamespaceUri { get; set; }
        public opcGatewayOption opcGatewayOptions { get; set; }
        public List<OPCNodeConfig> opcNodeConfig { get; set; }
    }
    public class OPCNode
    {
        public string name { get; set; }
        public string type { get; set; }
        public string template { get; set; }
        public string description { get; set; }
        public List<OPCNode> childTypes { get; set; }
        public List<string> ObjectLevels { get; set; } = new List<string>();
    }
    public class opcGatewayOption
    {
        public string saveDataItemsEdge { get; set; }
        public string saveDataItemsInterval { get; set; }
        public string saveDataItemsNode { get; set; }
    }
    public class OPCNodeConfig
    {
        [JsonProperty(Order = 1)]
        public string description { get; set; }
        [JsonProperty(Order = 2)]
        public string version { get; set; }
        [JsonProperty(Order = 3)]
        public string nameSpaceUri { get; set; }
        [JsonProperty(Order = 5)]
        public string xmlConfig { get; set; }
        [JsonProperty(Order = 4)]
        public List<OPCNode> opcNodes { get; set; }
    }
    public class NodeType
    {
        public bool IsObjectLevel { get; set; }
        public bool IsSimpleType { get; set; }
        public bool IsComplexType { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string SimpleType { get; set; }
        public string Template { get; set; }
        public List<string> ComplexTypes { get; set; }
    }
    public class OPCPage
    {
        public string PageName { get; set; }
        public List<OPCColumnType> objects { get; set; } = new List<OPCColumnType>();
    }
    public class OPCColumnType
    {
        public List<string> objectLevel { get; set; }
        public string tagName { get; set; }
        public string dataType { get; set; }
        public string description { get; set; }
    }

    /////////////////////////////////////////////////////PLC Type//////////////////////////////////////////////////////

    public class PLCDictionaryDataType
    {
        public string name { get; set; }
        public string version { get; set; }
        public string comment { get; set; }
        public List<OpcToPLCMappings> opcToPLCMappings { get; set; }
    }
    public class OpcToPLCMappings
    {
        public string plcType { get; set; }
        public string plcChannel { get; set; }
        public string plcModel { get; set; }
        public string plcAddress { get; set; }
        public int refreshRate { get; set; }
        public bool manualRead { get; set; }
        public int connectionTimeout { get; set; }
        public int transactionTimeout { get; set; }
        public int connectionAttempts { get; set; }
        public bool enabled { get; set; }
        public string plcSettings1 { get; set; }
        public string plcSettings2 { get; set; }
        public List<NodeMapping> nodeMapping { get; set; }
    }
    public class NodeMapping
    {
        public string plcTag { get; set; }
        public string plcTagType { get; set; }
        public string accessType { get; set; }
        public int plcTagElement { get; set; }
        public bool modifyValueBy10 { get; set; }
        public double modifyValueBy { get; set; }
        public string description { get; set; }
        public string template { get; set; }
        public string opcNode { get; set; }
    }
}

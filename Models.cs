using System;
using System.Collections.Generic;

namespace ConsoleApp1
{
    public class OpcUaWriteItem
    {
        public string NodeId { get; set; } = string.Empty;
        public object Value { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;

        public OpcUaWriteItem() { }

        public OpcUaWriteItem(string nodeId, object value, string description = "")
        {
            NodeId = nodeId;
            Value = value;
            Description = description;
        }
    }
}


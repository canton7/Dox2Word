using System;
using System.Collections.Concurrent;
using System.Xml.Serialization;

namespace Dox2Word.Parser
{
    public static class SerializerCache
    {
        private static readonly ConcurrentDictionary<(Type type, string? root), XmlSerializer> lookup = new();

        public static XmlSerializer Get<T>(string? root = null)
        {
            return lookup.GetOrAdd((typeof(T), root), key => 
                key.root == null ? new XmlSerializer(key.type) : new XmlSerializer(key.type, new XmlRootAttribute(key.root)));
        }
    }
}

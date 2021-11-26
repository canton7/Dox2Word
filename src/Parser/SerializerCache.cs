using System.Xml.Serialization;

namespace Dox2Word.Parser
{
    public static class SerializerCache<T>
    {
        public static readonly XmlSerializer Instance = new(typeof(T));
    }
}

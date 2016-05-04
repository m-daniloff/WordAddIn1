using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace Esquire.Common.Xml
{
    public static class Serializer
    {
        /// <summary>
        /// Deserializes XML, creating an object of Type t from an XElement serialization
        /// </summary>
        /// <param name="xElement">XElement serialization of object of Type t</param>
        /// <returns>Instantiated object of type T</returns>
        private static T CreateObjectFromXElement<T>(XElement xElement)
        {
            using (var memoryStream = new MemoryStream(Encoding.ASCII.GetBytes(xElement.ToString())))
            {
                var xmlSerializer = new XmlSerializer(typeof(T));
                return (T)xmlSerializer.Deserialize(memoryStream);
            }
        }

        /// <summary>
        /// Serializes an object of type T to XML
        /// </summary>
        /// <param name="obj">Object to serialize</param>
        /// <returns>XML representation</returns>
        public static XElement ObjectToXElement<T>(T obj)
        {
            using (var memoryStream = new MemoryStream())
            {
                using (TextWriter streamWriter = new StreamWriter(memoryStream))
                {
                    var xmlSerializer = new XmlSerializer(typeof(T));
                    xmlSerializer.Serialize(streamWriter, obj);
                    return XElement.Parse(Encoding.ASCII.GetString(memoryStream.ToArray()));
                }
            }
        }

        public static void SerializeDictionary<K, V>(XmlWriter writer, Dictionary<K, V> dictionary, string dictionaryName)
        {
            XElement node = SerializeDictionary<K, V>(dictionary, dictionaryName);
            node.WriteTo(writer);
        }

        public static XElement SerializeDictionary<K, V>(Dictionary<K, V> dictionary, string dictionaryName)
        {
            XElement dictionaryXml = new XElement(dictionaryName);
            foreach (KeyValuePair<K, V> keyVal in dictionary)
            {
                XElement keyValueNode = new XElement("KeyValue");
                XElement keyNode = new XElement("Key");
                keyNode.Add(ObjectToXElement<K>(keyVal.Key));
                keyValueNode.Add(keyNode);
                XElement valueNode = new XElement("Value");
                valueNode.Add(ObjectToXElement<V>(keyVal.Value));
                keyValueNode.Add(valueNode);
                dictionaryXml.Add(keyValueNode);
            }
            return dictionaryXml;
        }

        public static Dictionary<K, V> DeserializeDictionary<K, V>(XElement xmlDictionary)
        {
            Dictionary<K, V> dictionary = new Dictionary<K, V>();
            foreach (XElement fieldMapping in xmlDictionary.Elements("KeyValue"))
            {
                XElement keyNode = fieldMapping.Element("Key");
                K key = CreateObjectFromXElement<K>(keyNode.Elements().First());
                XElement valueNode = fieldMapping.Element("Value");
                V value = CreateObjectFromXElement<V>(valueNode.Elements().First());
                if (key != null && value != null)
                {
                    dictionary.Add(key, value);
                }
            }
            return dictionary;
        }

        public static void SerializeList<T>(XmlWriter writer, string listName, List<T> list)
        {
            XElement listRoot = SerializeList<T>(listName, list);
            listRoot.WriteTo(writer);
        }

        public static XElement SerializeList<T>(string listName, List<T> list)
        {
            XElement listRoot = new XElement(listName);
            foreach (T item in list)
            {
                listRoot.Add(ObjectToXElement<T>(item));
            }
            return listRoot;
        }

        public static List<T> DeserializeList<T>(XElement listNode)
        {
            List<T> list = new List<T>();
            foreach (var item in listNode.Elements())
            {
                list.Add(CreateObjectFromXElement<T>(item));
            }
            return list;
        }
    }
}

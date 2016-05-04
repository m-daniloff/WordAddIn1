using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace Esquire.Common.Xml
{
    /// <summary>
    /// Base class containing methods needed to serialize/deserialize objects of type T to/from 
    /// an XML repository stored in an xml file. 
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public abstract class XmlRepositoryBase<T>
    {
        protected XmlRepositoryBase()
        {
            _repositoryContents = new Dictionary<string, T>();
        }

        #region Repository Contents

        /// <summary>
        /// Full path of the repository XML file.
        /// </summary>
        protected abstract string RepositoryFilePath { get; }

        /// <summary>
        /// Name of the root element in the repository
        /// </summary>
        protected abstract string RepositoryNodeName { get; }

        /// <summary>
        /// Indicates if RepositoryContents has been populated with the contents of the XML file.
        /// </summary>
        private bool _repositoryLoaded;

        private Dictionary<string, T> _repositoryContents;
        /// <summary>
        /// Dictionary containing the contents of the repository indexed by each items name.
        /// </summary>
        protected Dictionary<string, T> RepositoryContents
        {
            get 
            {
                if (!_repositoryLoaded)
                {
                    LoadRepository();
                }
                return _repositoryContents; 
            }
        }

        /// <summary>
        /// Returns the name of a particular item.
        /// </summary>
        /// <param name="item">Item to identify</param>
        /// <returns>Items name</returns>
        protected abstract string GetItemId(T item);

        /// <summary>
        /// Returns the object identified by the name itemName
        /// </summary>
        /// <param name="itemId">Id of the item to return</param>
        /// <returns>Item referred to by itemName</returns>
        protected T GetItem(string itemId)
        {
            T item;
            RepositoryContents.TryGetValue(itemId, out item);
            return item;
        }

        /// <summary>
        /// Adds an item to the repository if it does not already exist, or updates the stored item 
        /// if it does.
        /// </summary>
        /// <param name="item">Item to add/update</param>
        protected void StoreItem(T item)
        {
            if (item == null)
            {
                throw new ArgumentException("item");
            }
            if (RepositoryContents.ContainsKey(GetItemId(item)))
            {
                RepositoryContents[GetItemId(item)] = item;
            }
            else
            {
                RepositoryContents.Add(GetItemId(item), item);
            }
            
        }

        /// <summary>
        /// Deletes the specified item from the repository
        /// </summary>
        /// <param name="itemId">Id of the item to delete.</param>
        protected void DeleteItem(string itemId)
        {
            RepositoryContents.Remove(itemId);
        }

        #endregion

        #region Save Repository

        /// <summary>
        /// Deserializes the contents of the repository to XML.
        /// </summary>
        /// <returns>XML representation of the repository</returns>
        protected XElement RepositoryToXml()
        {
            XElement repositoryRoot = new XElement(RepositoryNodeName);
            foreach (T repositoryElement in RepositoryContents.Values)
            {
                repositoryRoot.Add(ObjectToXElement(repositoryElement));
            }
            return repositoryRoot;
        }

        /// <summary>
        /// Saves the repository to the XML file specified by RepositoryFilePath
        /// </summary>
        protected virtual void SaveRepository()
        {            
            XElement repository = RepositoryToXml();
            repository.Save(RepositoryFilePath);
        }

        #endregion

        #region Load Repository

        /// <summary>
        /// Loads the contents of the repository from the xml file located at RepositoryFilePath.
        /// </summary>
        private void LoadRepository()
        {
            IEnumerable<XElement> repositoryNodes = GetRepositoryNodes(GetRootNode());
            foreach (XElement node in repositoryNodes)
            {
                T item = CreateObjectFromXElement(node);
                _repositoryContents.Add(GetItemId(item), item);
            }
            _repositoryLoaded = true;
        }

        /// <summary>
        /// Reads the xml file located at RepositoryFilePath and returns the root node.
        /// </summary>
        /// <returns>Root node or null if file does not exist</returns>
        protected virtual XElement GetRootNode()
        {
            XElement rootNode = null;
            if (File.Exists(RepositoryFilePath))
            {
                XDocument xDoc;
                using (FileStream xmlFile = new FileStream(RepositoryFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    xDoc = XDocument.Load(xmlFile);
                }
                rootNode = xDoc.Root;
            }
            return rootNode;
        }

        /// <summary>
        /// Returns all XML nodes which have the same node name as the name of type T.
        /// </summary>
        /// <param name="repositoryRoot">Repository root node</param>
        /// <returns>Enumerable collection of XElements each of which is a serialization of a object of type T</returns>
        private IEnumerable<XElement> GetRepositoryNodes(XElement repositoryRoot)
        {
            Type type = typeof(T);
            string nodeName = type.Name;
            return repositoryRoot != null ? repositoryRoot.Elements(nodeName) : new List<XElement>();
        }

        #endregion

        #region Item Serialization and Deserialization

        /// <summary>
        /// Deserializes XML, creating an object of Type t from an XElement serialization
        /// </summary>
        /// <param name="xElement">XElement serialization of object of Type t</param>
        /// <returns>Instantiated object of type T</returns>
        private static T CreateObjectFromXElement(XElement xElement)
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
        private XElement ObjectToXElement(T obj)
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

        #endregion
    }
}

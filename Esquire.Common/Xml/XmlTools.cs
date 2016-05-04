using System;
using System.Xml.Linq;

namespace Esquire.Common.Xml
{
    public class XmlTools
    {
        /// <summary>
        /// Gets the value of the specified child node from the XElement. If the node does not have the specified child the empty string 
        /// is returned.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="nodeName"></param>
        /// <returns></returns>
        public static string GetChildNodeValue(XElement element, string nodeName)
        {
            return GetChildNodeValue(element, nodeName, "");
        }

        /// <summary>
        /// Gets the value of the specified child node from the XElement. If the node does not have the specified child the specified
        /// default string is returned
        /// </summary>
        /// <param name="element"></param>
        /// <param name="nodeName"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static string GetChildNodeValue(XElement element, string nodeName, string defaultValue)
        {
            XElement childNode = element != null ? element.Element(nodeName) : null;
            return childNode != null ? childNode.Value : defaultValue;
        }

        /// <summary>
        /// Gets the value of the specified attribute from the XElement.
        /// </summary>
        /// <param name="element">XElement that has an attribute named attributeName</param>
        /// <param name="attributeName">Name of the attribute to get the value of</param>
        /// <param name="defaultValue">The default value to be returned if the specified attribute does not exist</param>
        /// <returns>The value of the attribure or the specified defaultvalue if the attribute is not found</returns>
        public static string GetAttributeValue(XElement element, string attributeName, string defaultValue)
        {
            XAttribute attribute = element != null ? element.Attribute(attributeName) : null;
            return attribute != null ? attribute.Value : defaultValue;
        }

        /// <summary>
        /// Gets the value of the specified attribute from the XElement.
        /// </summary>
        /// <param name="element">XElement that has an attribute named attributeName</param>
        /// <param name="attributeName">Name of the attribute to get the value of</param>
        /// <returns>The value of the attribure or an empty string if the attribute is not found</returns>
        public static string GetAttributeValue(XElement element, string attributeName)
        {
            return GetAttributeValue(element, attributeName, "");
        }

        /// <summary>
        /// Converts a string representation of a bool to a bool.
        /// Strings "1" and "true" get evaluate to true and strings "0" and "false" get evaluate to false
        /// </summary>
        /// <param name="value">String to be parsed as a bool</param>
        /// <param name="defaultValue">Default value to use of parsing the string fails</param>
        /// <returns>The result of parsing the string - or defaultValue if the string was no successfully parsed.</returns>
        public static bool ParseBool(string value, bool defaultValue)
        {
            bool result = defaultValue;
            if (!String.IsNullOrEmpty(value))
            {
                if (value == "1" || value.ToLower() == "true")
                {
                    result = true;
                }
                else if (value == "0" || value.ToLower() == "false")
                {
                    result = false;
                }
            }
            return result;
        }

        /// <summary>
        /// Converts a string representation of an interger to an int.
        /// </summary>
        /// <param name="str">String to be parsed as an integer</param>
        /// <param name="defaultValue"></param>
        /// <returns>The result of parsing the string - or defaultValue if the string was no successfully parsed.</returns>
        public static int ParseInteger(string str, int defaultValue)
        {
            int value;
            return int.TryParse(str, out value) ? value : defaultValue;
        }
    }
}

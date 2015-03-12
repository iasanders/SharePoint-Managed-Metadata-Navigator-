using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Xml.Linq;
using Microsoft.SharePoint.Taxonomy;

namespace MMDNavigator
{
    public class GenUtil
    {

        /// <summary>
        /// </summary>
        public static void LogIt(params object[] objs)
        {
            string output = "";

            for (int i = 0; i < objs.Length; i++)
            {
                if (objs[i] == null) objs[i] = "";

                string delim = " : ";

                if (i == objs.Length - 1) delim = "";

                output = string.Concat(output, objs[i], delim);
            }

            Debug.WriteLine(string.Format("[bandr] {0}", output));
        }


        /// <summary>
        /// </summary>
        public static string LabelsToString(LabelCollection labels, string termName)
        {
            if (labels == null || labels.Count <= 0)
                return "";

            List<string> lstLabels = labels.Where(x => x.Value != termName).Select(x => x.Value).ToList<string>();

            string strLabels = string.Join(";", lstLabels.ToArray<string>());

            return strLabels;
        }


        /// <summary>
        /// </summary>
        public static string SafeTrim(object o)
        {
            if (o == null)
                return "";
            else
                return o.ToString().Trim();
        }


        /// <summary>
        /// </summary>
        public static bool SafeToBool(object o)
        {
            if (IsNull(o))
                return false;
            else
            {
                bool result;
                if (!bool.TryParse(o.ToString(), out result))
                {
                    if (o.ToString() == "1" || o.ToString().Trim().ToLower() == "yes" || o.ToString().Trim().ToLower() == "true")
                        return true;
                    else
                        return false;
                }
                else
                    return result;
            }
        }


        /// <summary>
        /// Parse guid, if not a guid then return new guid.
        /// </summary>
        public static Guid SafeToGuid(object o)
        {
            if (!IsGuid(o))
            {
                return Guid.NewGuid();
            }

            return (new Guid(o.ToString()));
        }

        /// <summary>
        /// Convert to string, default is string.empty.
        /// </summary>
        public static string SafeXmlAttributeToString(XElement termElement, string attr)
        {
            if (termElement == null)
                return string.Empty;

            var xAttribute = termElement.Attribute(attr);

            return xAttribute != null ? xAttribute.Value.Trim() : string.Empty;
        }

        /// <summary>
        /// Conver to bool, default is false.
        /// </summary>
        public static bool SafeXmlAttributeToBool(XElement termElement, string attr)
        {
            if (termElement == null)
                return false;

            var xAttribute = termElement.Attribute(attr);

            return xAttribute == null ? false : SafeToBool(xAttribute.Value);
        }

        /// <summary>
        /// Convert to guid, default is new guid.
        /// </summary>
        public static Guid SafeXmlAttributeToGuid(XElement termElement, string attr)
        {
            if (termElement == null)
                return Guid.NewGuid();

            var xAttribute = termElement.Attribute(attr);

            return xAttribute == null ? Guid.NewGuid() : SafeToGuid(xAttribute.Value);
        }

        /// <summary>
        /// Convert to guid, default is new guid.
        /// </summary>
        public static Guid? SafeXmlAttributeToGuidOrNull(XElement termElement, string attr)
        {
            if (termElement == null)
                return null;

            var xAttribute = termElement.Attribute(attr);

            if (xAttribute == null)
                return null;

            if (!IsGuid(xAttribute.Value))
                return null;

            return SafeToGuid(xAttribute.Value);
        }


        /// <summary>
        /// </summary>
        public static bool IsGuid(object o)
        {
            if (IsNull(o))
            {
                return false;
            }

            try
            {
                Guid g = (new Guid(o.ToString()));
                return true;
            }
            catch(Exception exc)
            {
                return false;
            }
        }


        /// <summary>
        /// </summary>
        public static bool IsNull(object x)
        {
            if ((x == null)
                || (Convert.IsDBNull(x))
                || x.ToString().Trim().Length == 0)
                return true;
            else
                return false;
        }


        /// <summary>
        /// </summary>
        public static string CSVer(string s)
        {
            s = GenUtil.SafeTrim(s);

            if (s.Contains(","))
            {
                if (s.Contains("\""))
                    return string.Concat("\"", s.Replace("\"", "'"), "\"");
                else
                    return string.Concat("\"", s, "\"");
            }
            else
            {
                return s.Replace("\"", "'");
            }
        }


    }
}

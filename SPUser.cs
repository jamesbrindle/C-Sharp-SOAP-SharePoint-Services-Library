using System;
using System.Collections.Generic;
using System.Xml;

namespace SharePointIntegration.Objects
{
    /// <summary>
    /// A SharePoint user object
    /// </summary>
    public class SPUser
    {
        public int UserID { get; set; }
        public String Name { get; set; }
        public String LoginName { get; set; }
        public String Email { get; set; }
        public String Sid { get; set; }
        public String Notes { get; set; }
        public Boolean SiteAdministrator { get; set; }
        public Boolean DomainGroup { get; set; }
    }

    /// <summary>
    /// An object contains a collection (list) of SharePoint users. It also contains a method to deserialise a collections of users in the
    /// form of the XML document envelope to a SharePoint user collection object
    /// </summary>
    public class SPUserCollection
    {
        private List<SPUser> m_userColleciton = new List<SPUser>();
        public List<SPUser> UserList
        {
            get { return m_userColleciton; }
            set { m_userColleciton = value; }
        }

        /// <summary>
        /// Deserialises a collection of SharePoint users in the form of an XML document envelope into a SPUserCollection object
        /// </summary>
        /// <param name="xmlDocument">The XML document envelope to deserialise</param>
        /// <returns>The deserialised SPUserCollection object </returns>
        public static SPUserCollection ConvertToUserCollection(XmlDocument xmlDocument)
        {
            SPUserCollection col = new SPUserCollection();

            try
            {
                //soap:Body/GetAllUserCollectionFromWebResponse/GetAllUserCollectionFromWebResult/Users
                if (xmlDocument.DocumentElement.ChildNodes[0].ChildNodes[0].ChildNodes[0].ChildNodes[0] != null)
                {
                    foreach (XmlNode node in xmlDocument.DocumentElement.ChildNodes[0].ChildNodes[0].ChildNodes[0].ChildNodes[0])
                    {
                        if (node.Name == "Users")
                        {
                            for (int f = 0; f < node.ChildNodes.Count; f++)
                            {
                                SPUser usr = new SPUser();

                                usr.UserID = Convert.ToInt32(node.ChildNodes[f].Attributes["ID"].Value);
                                usr.Name = node.ChildNodes[f].Attributes["Name"].Value;
                                usr.LoginName = node.ChildNodes[f].Attributes["LoginName"].Value;
                                usr.Email = node.ChildNodes[f].Attributes["Email"].Value;
                                usr.Sid = node.ChildNodes[f].Attributes["Sid"].Value;
                                usr.Notes = node.ChildNodes[f].Attributes["Notes"].Value;
                                usr.SiteAdministrator = Convert.ToBoolean(node.ChildNodes[f].Attributes["IsSiteAdmin"].Value);
                                usr.DomainGroup = Convert.ToBoolean(node.ChildNodes[f].Attributes["IsDomainGroup"].Value);

                                col.UserList.Add(usr);
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.Out.WriteLine("Error (User): " + e.Message);
            }

            return col;
        }
    }
}

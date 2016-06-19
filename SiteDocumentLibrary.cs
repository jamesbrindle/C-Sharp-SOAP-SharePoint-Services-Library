using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace PowershareAdminConsole.SharePointObjects
{
    /// <summary>
    /// SharePoint document library object
    /// </summary>
    public class SiteDocumentLibrary
    {
        public String Path { get; set; }
        public String Title { get; set; }

        public SiteDocumentLibrary(String path, String title)
        {
            this.Path = path;
            this.Title = title;
        }
    }

    /// <summary>
    /// Object containing a collection (list) of SiteDocumentLibrary objects and a method to deserialise a collection of document
    /// libaries from a json string
    /// </summary>
    public class SiteDocumentLibraryCollection
    {
        private List<SiteDocumentLibrary> m_site_document_libs = new List<SiteDocumentLibrary>();

        public List<SiteDocumentLibrary> SiteDocumentLibraries
        {
            get { return m_site_document_libs; }
            set { m_site_document_libs = value; }
        }

        /// <summary>
        /// Deserialises a collection of document libraries in the form of a json string to a SiteDocumentLibaryCollection object
        /// </summary>
        /// <param name="jsonString">The json string to deserialse</param>
        /// <returns>The SiteDocumentLibraryCollection object</returns>
        public static SiteDocumentLibraryCollection GetSiteDocumentLibrariesFromJSONString(string jsonString)
        {
            SiteDocumentLibraryCollection col = new SiteDocumentLibraryCollection();

            if (!string.IsNullOrEmpty(jsonString))
            {
                jsonString = jsonString.Replace("{", "").Replace("}", "").Replace("\"", "").Replace(":", "");

                string[] siteParts = Regex.Split(jsonString, "], ");

                for (int i = 0; i < siteParts.Length; i++)
                {
                    string[] siteDocParts = Regex.Split(siteParts[i], "\\[");
                    string[] docParts = Regex.Split(siteDocParts[1], ", ");

                    for (int j = 0; j < docParts.Length; j++)
                    {
                        SiteDocumentLibrary lib = new SiteDocumentLibrary(siteDocParts[0].Replace(" ", ""), 
                            j == docParts.Length - 1 && i == siteParts.Length - 1 ? docParts[j].Substring(0, docParts[j].Length - 1) : docParts[j]);

                        col.SiteDocumentLibraries.Add(lib);
                    }                    
                }                
            }

            return col;
        }

        /// <summary>
        /// Serialises a SiteDocumentLibrary collection object to a valid json string
        /// </summary>
        /// <param name="col">The SiteDocumentLibraryCollection object to serialise</param>
        /// <returns>The serialised json string of site document libraries</returns>
        public static String GetDocumentLibrariesJSONStringFromCollection(SiteDocumentLibraryCollection col)
        {
            string jsonString = "";

            if (col != null)
            {
                if (col.SiteDocumentLibraries.Count > 0)
                {
                    jsonString += "{";

                    List<string> paths = new List<string>();

                    // Create a list of path names to iterate through
                    for (int i = 0; i < col.SiteDocumentLibraries.Count; i++)
                    {
                        if (!paths.Contains("\"" + col.SiteDocumentLibraries[i].Path + "\": "))
                            paths.Add(col.SiteDocumentLibraries[i].Path == "SiteRoot" ? "\"\": " : "\"" + col.SiteDocumentLibraries[i].Path + "\": ");
                    }

                    for (int i = 0; i < paths.Count; i++)
                    {
                        jsonString += paths[i] + "[";

                        List<string> libTitles = new List<string>();

                        // create a list of names specific to a particular path
                        for (int j = 0; j < col.SiteDocumentLibraries.Count; j++)
                        {
                            string pathToMatch = "\"" + col.SiteDocumentLibraries[j].Path + "\": ";

                            if (pathToMatch == paths[i])
                                libTitles.Add("\"" + col.SiteDocumentLibraries[j].Title + "\"");
                        }

                        // create the json string, a combination of the path names followed by a list of library names in square brackets
                        for (int j = 0; j < libTitles.Count; j++)
                            jsonString += j == libTitles.Count - 1 ? libTitles[j] : libTitles[j] + ", ";

                        jsonString += i == paths.Count - 1 ? "]" : "], ";
                    }

                    jsonString += "}";                  
                }
            }

            return jsonString;
        }

        /// <summary>
        /// Method to determine of a document library within a document library collection contains a particular 
        /// SharePoint library title and site path
        /// </summary>
        /// <param name="path">The matching path to search for</param>
        /// <param name="title">The matching library name to search for</param>
        /// <returns>True if the search found a matching path and library title, false otherwise</returns>
        public Boolean SiteDocumentCollectionContains(string path, string title)
        {
            foreach (SiteDocumentLibrary lib in SiteDocumentLibraries)
            {
                if (lib.Path == path && lib.Title == title)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Removes a document library from a document library collection given a particular library name and site path
        /// </summary>
        /// <param name="path">The matching path to search for deletion</param>
        /// <param name="title">The matching name to search for deletion</param>
        public void RemoveDocumentLibrary(string path, string title)
        {
            for (int i = 0; i < SiteDocumentLibraries.Count; i++)
            {
                if (SiteDocumentLibraries[i].Path == path && SiteDocumentLibraries[i].Title == title)
                    SiteDocumentLibraries.RemoveAt(i);
            }
        }
    }
}

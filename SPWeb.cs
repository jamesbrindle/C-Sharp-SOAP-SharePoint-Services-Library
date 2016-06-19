using System;

namespace PowershareAdminConsole.SharePointObjects
{
    /// <summary>
    /// A SharePoint site object
    /// </summary>
    public class SPWeb
    {
        public String Title { get; set; }
        public String Url { get; set; }
        public String Path { get; set; }

        public SPWeb()
        { }

        public SPWeb(string title, string Url, string path)
        {
            this.Title = title;
            this.Url = Url;
            this.Path = path;
        }
    }
}

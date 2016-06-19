using System;

namespace PowershareAdminConsole.SharePointObjects
{
    /// <summary>
    /// A SharePoint library object
    /// </summary>
    public class SPList
    {
        public enum BaseTypeEnum
        {
            None,
            GenericList,
            DocumentLibrary,
            Unused,
            DiscussionBoard,
            Survey,
            Issue
        }

        public String DefaultViewUrl { get; set; }
        public String ID { get; set; }
        public BaseTypeEnum BaseType { get; set; }
        public String Host { get; set; }
        public String Name { get; set; }
        public String Path { get; set; }
        public String Title { get; set; }

        public SPList(String host)
        {
            Host = host;
        }

        public SPList(string host, string defaultViewUrl, string id, int baseType, string path, string title)
        {
            DefaultViewUrl = defaultViewUrl;
            ID = id;
            Host = host;
            Path = path;
            Title = title;

            switch (baseType)
            {
                case (-1):
                    BaseType = BaseTypeEnum.None;
                    break;
                case (0):
                    BaseType = BaseTypeEnum.GenericList;
                    break;
                case (1):
                    BaseType = BaseTypeEnum.DocumentLibrary;
                    break;
                case (2):
                    BaseType = BaseTypeEnum.Unused;
                    break;
                case (3):
                    BaseType = BaseTypeEnum.DiscussionBoard;
                    break;
                case (4):
                    BaseType = BaseTypeEnum.Survey;
                    break;
                case (5):
                    BaseType = BaseTypeEnum.Issue;
                    break;
                default:
                    BaseType = BaseTypeEnum.None;
                    break;
            }
        }
    }
}

using System;
using System.IO;

namespace PowershareAdminConsole.SharePointObjects
{
    /// <summary>
    /// A collection of SharePoint XML envelopes to be sent to the SharePoint server. Some envelopes may require certain perameters and the XML 
    /// envelopes then generated and returned
    /// </summary>
    public class SPEnvelope
    {
        /// <summary>
        /// A collection of SharePoint sites
        /// </summary>
        /// <returns></returns>
        public static String GetSubWebCollectionEnvelope()
        {
            return
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">" +
                    "<soap:Body>" +
                        "<GetAllSubWebCollection xmlns=\"http://schemas.microsoft.com/sharepoint/soap/\" />" +
                    "</soap:Body>" +
                "</soap:Envelope>";
        }

        /// <summary>
        /// A collection of SharePoint libraries
        /// </summary>
        /// <returns></returns>
        public static String GetListCollectionEnvelope()
        {
            return
                 "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">" +
                    "<soap:Body>" +
                        "<GetListCollection xmlns=\"http://schemas.microsoft.com/sharepoint/soap/\" />" +
                    "</soap:Body>" +
                "</soap:Envelope>";
        }

        /// <summary>
        /// A collection of documents contained within a document library
        /// </summary>
        /// <param name="listNameOrGuid">The document library of where to retrieve the documents from</param>
        /// <returns></returns>
        public static String GetDocumentCollectionEnvelope(string listNameOrGuid)
        {
            return
                "<?xml version=\"1.0\" encoding=\"utf-8\" ?>" +
                "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">" +
                    "<soap:Body>" +
                        "<GetListItems xmlns=\"http://schemas.microsoft.com/sharepoint/soap/\">" +
                            "<listName>" + listNameOrGuid + "</listName>" +
                            "<query>" +
                                "<Query xmlns=\"\">" +
                                    "<Where>" +
                                        "<Gt>" +
                                            "<FieldRef Name=\"ID\" />" +
                                            "<Value Type=\"Number\">0</Value>" +
                                        "</Gt>" +
                                    "</Where>" +
                                "</Query>" +
                            "</query>" +
                            "<viewFields>" +
                                "<ViewFields xmlns=\"\">" +
                                    "<FieldRef Name=\"UniqueId\" />" +
                                    "<FieldRef Name=\"Title\" />" +
                                    "<FieldRef Name=\"FileRef\" />" +
                                    "<FieldRef Name=\"Modified\" />" +
                                    "<FieldRef Name=\"Created\" />" +
                                    "<FieldRef Name=\"Editor\" />" +
                                "</ViewFields>" +
                            "</viewFields>" +
                           "<queryOptions>" +
                                "<QueryOptions xmlns=\"\">" +
                                    "<IncludeMandatoryColumns>True</IncludeMandatoryColumns>" +
                                    "<DateInUtc>TRUE</DateInUtc>" +
                                "</QueryOptions>" +
                            "</queryOptions>" +
                        "</GetListItems>" +
                    "</soap:Body>" +
                "</soap:Envelope>";
        }

        /// <summary>
        /// A collection of documents, which have had changes since a given date
        /// </summary>
        /// <param name="listNameOrGuid">The document library in which to retrieve the documents that have changes</param>
        /// <param name="fromDate">The date in which to retrieve the document changes from</param>
        /// <returns></returns>
        public static String GetDocumentChangesCollectionEnvelope(string listNameOrGuid, DateTime fromDate)
        {
            return
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">" +
                    "<soap:Body>" +
                        "<GetListItemChanges xmlns=\"http://schemas.microsoft.com/sharepoint/soap/\">" +
                            "<listName>" + listNameOrGuid + "</listName>" +
                             "<viewFields>" +
                                "<ViewFields xmlns=\"\">" +
                                    "<FieldRef Name=\"UniqueId\" />" +
                                    "<FieldRef Name=\"Title\" />" +
                                    "<FieldRef Name=\"FileRef\" />" +
                                    "<FieldRef Name=\"Modified\" />" +
                                    "<FieldRef Name=\"Created\" />" +
                                    "<FieldRef Name=\"Editor\" />" +
                                "</ViewFields>" +
                            "</viewFields>" +
                            "<since>" + fromDate.ToLongDateString() + "</since>" +
                        "</GetListItemChanges>" +
                    "</soap:Body>" +
                "</soap:Envelope>";
        }

        /// <summary>
        /// A list of documents which are included in a given change token
        /// </summary>
        /// <param name="listNameOrGui">The document library in which to retrieve the documents which have had changes</param>
        /// <param name="changeTokenId">The change token ID (a unique ID, created when a list of changes is retrieved so we know which 
        /// documents have already been called for</param>
        /// <returns></returns>
        public static String GetDocumentChangesCollectionSinceTokenEnvelope(string listNameOrGui, string changeTokenId)
        {
            return
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">" +
                    "<soap:Body>" +
                        "<GetListItemChangesSinceToken xmlns=\"http://schemas.microsoft.com/sharepoint/soap/\">" +
                            "<listName>" + listNameOrGui + "</listName>" +
                            "<query>" +
                                "<Query xmlns=\"\">" +
                                    "<Where>" +
                                        "<Gt>" +
                                            "<FieldRef Name=\"ID\" />" +
                                            "<Value Type=\"Number\">0</Value>" +
                                        "</Gt>" +
                                    "</Where>" +
                                "</Query>" +
                            "</query>" +
                            "<viewFields>" +
                                "<ViewFields xmlns=\"\">" +
                                    "<FieldRef Name=\"UniqueId\" />" +
                                    "<FieldRef Name=\"Title\" />" +
                                    "<FieldRef Name=\"FileRef\" />" +
                                    "<FieldRef Name=\"Modified\" />" +
                                    "<FieldRef Name=\"Created\" />" +
                                    "<FieldRef Name=\"Editor\" />" +
                                "</ViewFields>" +
                            "</viewFields>" +
                            "<queryOptions>" +
                                "<QueryOptions xmlns=\"\">" +
                                    "<IncludeMandatoryColumns>True</IncludeMandatoryColumns>" +
                                    "<DateInUtc>TRUE</DateInUtc>" +
                                "</QueryOptions>" +
                            "</queryOptions>" +
                            (String.IsNullOrEmpty(changeTokenId) ? "" : "<changeToken>" + changeTokenId + "</changeToken>") +
                        "</GetListItemChangesSinceToken>" +
                    "</soap:Body>" +
                "</soap:Envelope>";
        }

        /// <summary>
        /// The current version of a document given a particular document filename (the filename is a URL including SharePoint site)
        /// </summary>
        /// <param name="fileName">The filename of the document to retrieve the version for</param>
        /// <returns></returns>
        public static String GetDocumentVersionEnvelope(string fileName)
        {
            return
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">" +
                    "<soap:Body>" +
                        "<GetVersions xmlns=\"http://schemas.microsoft.com/sharepoint/soap/\">" +
                            "<fileName>" + fileName + "</fileName>" +
                        "</GetVersions>" +
                    "</soap:Body>" +
                "</soap:Envelope>";
        }

        /// <summary>
        /// Envelope used to upload a document, given the local document location and the target upload location
        /// </summary>
        /// <param name="filePathToUpload">The location including filename of the document to upload</param>
        /// <param name="host">The SharePoint domain URL</param>
        /// <param name="uploadLocation">The site / library location of where to upload the document to</param>
        /// <returns></returns>
        public static String GetDocumentUploadEnvelope(string filePathToUpload, string host, string uploadLocation)
        {
            return
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">" +
                "<soap:Body>" +
                    "<CopyIntoItems xmlns=\"http://schemas.microsoft.com/sharepoint/soap/\">" +
                        "<SourceUrl>" + filePathToUpload + "</SourceUrl>" +
                            "<DestinationUrls>" +
                                "<string>" + host + uploadLocation + "/" + Path.GetFileName(filePathToUpload) + "</string>" +
                            "</DestinationUrls>" +
                            "<Fields>" +
                                "<FieldInformation Type=\"File\" DisplayName=\"" +
                                Path.GetFileName(filePathToUpload) + "\" InternalName=\"" +
                                Path.GetFileName(filePathToUpload) + "\" Value=\"" + Path.GetFileName(filePathToUpload) + "\" />" +
                            "</Fields>" +
                        "<Stream>" + Convert.ToBase64String(File.ReadAllBytes(filePathToUpload)) + "</Stream>" +
                    "</CopyIntoItems>" +
                "</soap:Body>" +
            "</soap:Envelope>";
        }

        /// <summary>
        /// Permissions of a site / library / document
        /// </summary>
        /// <param name="objectName">The name of the site / library / document to check permissions for</param>
        /// <param name="objectType">The type of object to check permissions for, i.e. site / library / document</param>
        /// <returns></returns>
        public static String GetPermissionCollectionEnvelope(string objectName, string objectType)
        {
            return
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">" +
                    "<soap:Body>" +
                        "<GetPermissionCollection xmlns=\"http://schemas.microsoft.com/sharepoint/soap/directory/\">" +
                            "<objectName>" + objectName + "</objectName>" +
                            "<objectType>" + objectType + "</objectType>" +
                        "</GetPermissionCollection>" +
                    "</soap:Body>" +
                "</soap:Envelope>";
        }

        /// <summary>
        /// A list of permissions of the users for a given SharePoint site
        /// </summary>
        /// <returns></returns>
        public static String GetRolesAndPermissionsForSiteEnvelope()
        {
            return
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">" +
                    "<soap:Body>" +
                        "<GetRolesAndPermissionsForSite xmlns=\"http://schemas.microsoft.com/sharepoint/soap/directory/\" />" +
                    "</soap:Body>" +
                "</soap:Envelope>";
        }

        /// <summary>
        /// A list of users from a given SharePoint site
        /// </summary>
        /// <returns></returns>
        public static String GetAllUserCollectionFromWeb()
        {
            return
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">" +
                    "<soap:Body>" +
                        "<GetAllUserCollectionFromWeb xmlns=\"http://schemas.microsoft.com/sharepoint/soap/directory/\" />" +
                    "</soap:Body>" +
                "</soap:Envelope>";
        }
    }
}

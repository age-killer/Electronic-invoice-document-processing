using System;
using System.Net;
using System.Collections;
using Independentsoft.Exchange;

namespace Sample
{
    class Program
    {
        private static Hashtable folderTable = new Hashtable();
        private static Service service = null;

        static void Main(string[] args)
        {
            NetworkCredential credential = new NetworkCredential("username", "password");
            service = new Service("https://myserver/ews/Exchange.asmx", credential);

            try
            {
                Folder publicFolderRoot = service.GetFolder(StandardFolder.PublicFoldersRoot);

                SearchPublicFolder(publicFolderRoot.FolderId);

                foreach (string folderId in folderTable.Keys)
                {
                    Folder folder = (Folder)folderTable[folderId];

                    string folderPath = GetFolderPath(folder);

                    Console.WriteLine(folder.DisplayName);
                    Console.WriteLine(folder.FolderClass);
                    Console.WriteLine(folder.CreationTime);
                    Console.WriteLine(folderPath);
                    Console.WriteLine("--------------------------------------------------------------------");
                }

                Console.Read();
            }
            catch (ServiceRequestException ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.WriteLine("Error: " + ex.XmlMessage);
                Console.Read();
            }
            catch (WebException ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.Read();
            }
        }

        public static void SearchPublicFolder(FolderId folderId)
        {
            FindFolderResponse findFolderResponse = service.FindFolder(folderId, FolderPropertyPath.AllPropertyPaths);

            for (int i = 0; i < findFolderResponse.Folders.Count; i++)
            {
                Folder folder = findFolderResponse.Folders[i];

                folderTable.Add(folder.FolderId.Id, folder);

                if (folder.HasSubFolders)
                {
                    SearchPublicFolder(folder.FolderId);
                }
            }
        }

        static string GetFolderPath(Folder folder)
        {
            string path = "";

            path = path + "/" + folder.DisplayName;

            string parentId = folder.ParentId.Id;

            if (folderTable[parentId] != null)
            {
                Folder parentFolder = (Folder)folderTable[parentId];

                return GetFolderPath(parentFolder) + path;
            }

            return path;
        }
    }
}
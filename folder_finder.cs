using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using flookup_exception;
using System.Configuration;

namespace flookup
{
    public class FolderFinder
    {

        /// <summary>
        /// Adds each path in a folder to a string
        /// </summary>
        /// <param name="folder"> a folder being searched through</param>
        /// <returns> a string containing all the folders and sub folders in a root outlook folder</returns>
        public string EnumerateFolders(Outlook.Folder folder)
        {
            string ret = "";
            Outlook.Folders childFolders = folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Outlook.Folder childFolder in childFolders)
                {
                    // Write the folder path.
                    //Console.WriteLine(childFolder.FolderPath);
                    // Optimization possibly: consider only adding correct folder path to string

                    ret += childFolder.FolderPath + "\n";
                    // Call EnumerateFolders using childFolder and add result to ret.
                    ret += EnumerateFolders(childFolder);
                }
            }
            return ret;

        }


        /// <summary>
        ///  Retrieves all the stores tied to the outlook account currently opened
        /// </summary>
        /// <returns> a List containing all the inboxes in an outlook account</returns>
        public List<Outlook.Store> getStores()
        {

            List<Outlook.Store> stores = new List<Outlook.Store>();

            var appOutlook = new Microsoft.Office.Interop.Outlook.Application();
            // Get each store in the outlook account
            var outlookStores = appOutlook.Session.Stores;
            // add each store to a list (can most likely forgoe this and just return outlookStores
            foreach (Outlook.Store store in outlookStores)
            {
                stores.Add(store);
            }

            return stores;
        }
        /// <summary>
        /// Gets the inbox name associated with each store
        /// </summary>
        /// <param name="stores"> A list of stores for a paritcular outlook account</param>
        /// <returns> a string containing each inbox name for the currently opened outlook
        ///             account</returns>
        public string getAccounts(List<Outlook.Store> stores)
        {
            string ret = "";
            foreach (Outlook.Store store in stores)
            {
                ret += store.DisplayName + " \n";

            }
            return ret;
        }

        /// <summary>
        /// Given a list of stores this will rerieve a specified inbox
        /// </summary>
        /// <param name="inboxName"> the inbox being looked for</param>
        /// <param name="stores"> a list of stores tied to the opened outlook account</param>
        /// <returns></returns>
        public Outlook.Store findInbox(string inbox, List<Outlook.Store> stores)
        {
            foreach (Outlook.Store store in stores)
            {
                if (store.DisplayName.Equals(inbox))
                {
                    return store;
                }
            }
            // If the search did not return a result throw an InvalidInboxException
            throw new InvalidInboxException("Inbox " + inbox + " was not found.");

        }

        /// <summary>
        /// Finds any folder paths containing a particular folderName
        /// </summary>
        /// <param name="folderName"> the folder we are searching for</param>
        /// <param name="inbox"> the inbox we are looking at</param>
        public void findFolder(string folderName, Outlook.Store inbox)
        {
            string paths = "";
            // Get the root folder
            Outlook.Folder root = inbox.GetRootFolder() as Outlook.Folder;

            // get results
            paths = this.EnumerateFolders(root);
            //Console.WriteLine(paths);

            // create a regex for the folderName
            MatchCollection mc = Regex.Matches(paths, ".*" + folderName + ".*");

            if (mc.Count.Equals(0))
            {
                throw new InvalidFolderException("Folder " + folderName + " could not be found.");
            }

            foreach (Match match in mc)
            {
                Console.WriteLine("\t" + match.ToString().Substring(2));
            }

        }
        /// <summary>
        /// Prints out a help message giving information on options and parameters
        /// </summary>
        public void help()
        {
            Console.WriteLine("-----------Help-----------");
            Console.WriteLine("Retrieve folder paths for default inbox");
            Console.WriteLine("flookup {folder_name}");
            Console.WriteLine("");
            Console.WriteLine("Retrieve folder paths for a specified inbox and folder");
            Console.WriteLine("flookup -i {inbox} {folder_name}");
            Console.WriteLine("");
            Console.WriteLine("Retrieve all stores(inboxes) for current outlook session");
            Console.WriteLine("folder_finder -s");
            Console.WriteLine("");
            Console.WriteLine("Get the default inbox");
            Console.WriteLine("folder_finder -d");
            Console.WriteLine("");
            Console.WriteLine("Set the default inbox (inbox must be valid)");
            Console.WriteLine("folder_finder -d {valid inbox}");
        }
        /// <summary>
        /// Sets the default inbox to be used when preforming an flookup
        /// </summary>
        /// <param name="inbox"> default inbox name </param>
        /// <param name="stores"> a list of each store for a particular outlook account</param>
        public void setDefaultEmail(string inbox, List<Outlook.Store> stores)
        {
            try
            {
                Outlook.Store store = this.findInbox(inbox, stores);
                Configuration config = ConfigurationManager.OpenExeConfiguration("");
                //ConfigurationManager.AppSettings.Settings["defaultStore"].Value = inbox;
                config.AppSettings.Settings.Remove("defaultStore");
                config.AppSettings.Settings.Add("defaultStore", inbox);

                config.Save(ConfigurationSaveMode.Modified);
            }
            catch(InvalidInboxException e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(this.getAccounts(stores));
            }
        }
        /// <summary>
        /// Gets whatever the user set the default flookup email too
        /// </summary>
        /// <returns> the default inbox as a string</returns>
        public string getDefaultEmail()
        {
            string defaultEmail = ConfigurationManager.AppSettings["defaultStore"];
            if (defaultEmail.Equals(""))
            {
                throw new InvalidInboxException("No default email setup");
            }
            return defaultEmail;


        }
        /// <summary>
        /// prints out all the folder names (used for testing)
        /// </summary>
        private void printFolders()
        {

            var appOutlook = new Microsoft.Office.Interop.Outlook.Application();
            var folders = appOutlook.Session.Folders;
            var stores = appOutlook.Session.Stores;
            Console.WriteLine(stores.GetType());


            foreach (Microsoft.Office.Interop.Outlook.Store store in stores)
            {
                Console.WriteLine(store.DisplayName);

            }

        }
    }

}

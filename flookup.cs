using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using flookup_exception;
using System.Collections.Generic;

namespace flookup
{
    /// <summary>
    /// This is the main class for executing a folder look up
    /// Compile the solution and the execute the exe as a command line operation
    /// </summary>
    class flookup
    {
        public static void Main(string[] args)
        {
            FolderFinder finder = new FolderFinder();

            // case where user supplys a folder name however no -s option
            // and specifies only a folder name
            if (args.Length.Equals(1))
            {
                // User option which retreives all the stores(inboxes) for a particular account
                if (args[0].Equals("-s"))
                {
                    var stores = finder.getStores();
                    Console.Write(finder.getAccounts(stores));
                }
                // User option which retreives the default mailbox
                else if (args[0].Equals("-d"))
                {
                    try
                    {
                        Console.WriteLine(finder.getDefaultEmail());
                    }
                    catch(InvalidInboxException e)
                    {
                        Console.WriteLine(e.Message);
                    }
                    
                }
                else
                {
                    try
                    {
                        // Retrieve all the stores in an account
                        var stores = finder.getStores();
                        // Search for the specified Store
                        // This is the default store
                        Outlook.Store inbox = finder.findInbox(finder.getDefaultEmail(), stores);
                        finder.findFolder(args[0], inbox);


                    }
                    catch (InvalidInboxException e)
                    {
                        Console.WriteLine(e.Message);
                    }
                    catch (InvalidFolderException e)
                    {

                        Console.WriteLine(e.Message);
                    }
                }
            }
            else if (args.Length == 2)
            {
                if (!args[0].Equals("-d"))
                {
                    finder.help();
                }
                else
                {
                    List<Outlook.Store> stores = finder.getStores();
                    finder.setDefaultEmail(args[1], stores);
                }
            }
            // User uses option -s and specifies an inbox followed by a folder name
            else if (args.Length == 3)
            {
                // i stands for inbox
                if (!args[0].Equals("-i"))
                {
                    Console.WriteLine("You've entered an invalid option");
                    finder.help();
                }
                else
                {
                    try
                    {
                        // Retrieve all the stores in an account
                        var stores = finder.getStores();
                        // Search for the specified Store
                        // This is the default store
                        Outlook.Store inbox = finder.findInbox(args[1], stores);
                        finder.findFolder(args[2], inbox);
                    }
                    catch (InvalidInboxException e)
                    {
                        Console.WriteLine(e.Message);
                    }
                    catch (InvalidFolderException e)
                    {
                        Console.WriteLine(e.Message);
                    }
                }
            }
            // Error message
            else
            {
                finder.help();

            }

        }
    }
}

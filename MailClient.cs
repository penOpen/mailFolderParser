using MailKit;
using MailKit.Net.Imap;
using MimeKit;

namespace mailFolderParser
{
    internal class MailClient
    {
        private readonly ImapClient? Client;

        public MailClient()
        {
            Client = new ImapClient();
        }


        public IMailFolder GetInbox() => Client.Inbox;

        public void ConnectToMail()
        {
            try
            {
                Client.Connect("imap.mail.ru", 993, true);
                Console.WriteLine("Client connection: " + Client.IsConnected);
            }
            catch (Exception)
            {
                Console.WriteLine("Error: Invalid server");
                Environment.Exit(-1);
            }
        }

        public void AuthentificateUser(string mail, string password)
        {
            try
            {
                Client.Authenticate(mail, password);
                Console.WriteLine("Client Authentification: " + Client.IsAuthenticated); ;
            }
            catch (Exception)
            {
                Console.WriteLine("Error: invalid mail or password");
                Environment.Exit(-1);
            }
        }

        public void CheckFolders(Dictionary<string, string> dictionary)
        {
            HashSet<string> folders = new HashSet<string>(dictionary.Values);
            var inbox = Client.GetFolder("INBOX");
            foreach (var folder in folders)
            {
                try
                {
                    //Trying to get a required subfloder, create on fail
                    if (inbox.GetSubfolder(folder) != null)
                    {
                        Console.WriteLine("Folder " + folder + " exists!");
                    }
                }
                catch (Exception)
                {
                    Console.WriteLine("Cant find folder " + folder);
                    var newFolder = Client.GetFolder("INBOX");
                    newFolder.Create(folder, false);
                    Console.WriteLine("Created folder " + folder);
                }
            }

            try
            {
                //Trying to get an Unsorted subfolder, create on fail 
                if (inbox.GetSubfolder("Unsorted") != null)
                {
                    Console.WriteLine("Folder Unsorted exists!");
                }
            }
            catch (Exception)
            {
                var newFolder = Client.GetFolder("INBOX");
                newFolder.Create("Unsorted", false);
                Console.WriteLine("Created folder Unsorted");
            }
        }

        public int MoveMessage(UniqueId u, MimeMessage msg, Dictionary<string, string> dictionary)
        {
            var inbox = Client.GetFolder("INBOX");
            var subject = msg.Subject.ToUpper();
            //Splitting subject of message into list
            List<string> wordsList = subject.Split(' ').ToList();
            foreach (var word in wordsList)
            {
                //Checking if dictionary contains word (need to fix)
                if (dictionary.ContainsKey(word))
                {
                    //Moving to destination folder
                    var destination = dictionary[word];
                    inbox.MoveTo(u, inbox.GetSubfolder(destination));
                    Console.WriteLine("Message [UID: " + u + "] MOVED TO " + destination);
                    return 1;
                }
                else
                {
                    //Moving to Unsorted folder
                    inbox.MoveTo(u, inbox.GetSubfolder("Unsorted"));
                    Console.WriteLine("Message [UID: " + u + "] MOVED TO Unsorted");
                    return 0;
                }
            }

            return -1;
        }

        public void DisconnectClient()
        {
            Client.Disconnect(true);
        }
    }
}

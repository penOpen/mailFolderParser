using mailFolderParser;
using MailKit;
using MailKit.Search;

//Init classes
var data = new UserData();
var client = new MailClient();

//Client connection and authentification
client.ConnectToMail();
client.AuthentificateUser(data.Mail, data.Password);

//Setting up inbox and uids
var _inbox  = client.GetInbox();
IMyMailFolder inbox = (IMyMailFolder)_inbox;
//var inbox = client.GetInbox();
var uids = inbox.Search(SearchQuery.All);

//Checking subfolders for existance
client.CheckFolders(data.Folders);

Console.WriteLine("Listening on inbox messages");

while (true)
{
    //5 second delay
    Thread.Sleep(5000);
    //Checking for new messages
    var newUids = inbox.Search(SearchQuery.Not(SearchQuery.Uids(uids)));
    if (newUids.Count > 0)
    {
        foreach (var u in newUids)
        {
            var msg = inbox.GetMessage(u);
            Console.WriteLine("New message from" + msg.From + " With subject: " + msg.Subject);
            uids.Add(u);
            //Sending message for sorting (need to correct condition to send message to Unsorted subfolder)
            int result = client.MoveMessage(u, msg, data.Folders);
        }
    }
    else
    {
        Console.WriteLine("Zero messages");
    }
    //On pressed key Q loop ends
    if (Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Q)
    {
        Console.WriteLine("Stop listening");
        break;
    }
}

client.DisconnectClient();
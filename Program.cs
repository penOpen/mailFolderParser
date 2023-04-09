using MailKit;
using MailKit.Search;

//Initializing a cancellation token for the wait time
var cancellationTokenSource = new CancellationTokenSource();
var cancellationToken = cancellationTokenSource.Token;

//Getting user mail, password
var info = Methods.GetUserInfo();
//Getting dictionary of folders where key -> sortCondition, value -> folder name
var dictionary = Methods.GetFolders();
//Connecting to IMAP mail server
var client = Methods.ConnectToMail();
//User authentication
client = Methods.AuthentificateUser(client, info);

//Setting up inbox and uids
var inbox = client.Inbox;
inbox.Open(FolderAccess.ReadWrite);
var uids = inbox.Search(SearchQuery.All);

//Checking subfolders for existance
Methods.CheckFolders(client, dictionary);

Console.WriteLine("Listening on inbox messages");

while (true)
{
    //Setting 5 second delay
    cancellationToken.WaitHandle.WaitOne(TimeSpan.FromSeconds(5));
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
            int result = Methods.MoveMessage(client, u, msg, dictionary);
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

client.Disconnect(true);
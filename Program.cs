using MailKit;
using MailKit.Search;


var cancellationTokenSource = new CancellationTokenSource();
var cancellationToken = cancellationTokenSource.Token;

var info = Methods.GetUserInfo();
var dictionary = Methods.GetFolders();

var client = Methods.ConnectToMail();
client = Methods.AuthentificateUser(client, info);

var inbox = client.Inbox;
inbox.Open(FolderAccess.ReadWrite);
var uids = inbox.Search(SearchQuery.All);

Methods.CheckFolders(client, dictionary);

Console.WriteLine("Listening on inbox messages");

while (true)
{
    cancellationToken.WaitHandle.WaitOne(TimeSpan.FromSeconds(5));
    var newUids = inbox.Search(SearchQuery.Not(SearchQuery.Uids(uids)));
    if (newUids.Count > 0)
    {
        foreach (var u in newUids)
        {
            var msg = inbox.GetMessage(u);
            Console.WriteLine("New message from" + msg.From + " With subject: " + msg.Subject);
            uids.Add(u);
            int result = Methods.MoveMessage(client, u, msg, dictionary);
        }
    }
    else
    {
        Console.WriteLine("Zero messages");
    }

    if (Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Q)
    {
        Console.WriteLine("Stop listening");
        break;
    }
}

client.Disconnect(true);
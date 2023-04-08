using System;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using ExcelDataReader;
using System.Data;

public class UserInfo
{
    public string? mail;
    public string? password;

    public UserInfo(string? mail, string? password)
    {
        this.mail = mail;
        this.password = password;
    }
}


public class Methods
{
    public static UserInfo GetUserInfo()
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        using var stream = File.Open(@"../../../data.xlsx", FileMode.Open, FileAccess.Read);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
        {
            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
            {
                UseHeaderRow = true
            }
        });
        var table = result.Tables[0];
        string? mail = table.Rows[0][1].ToString();
        string? password = table.Rows[1][1].ToString();
        stream.Close();
        var info = new UserInfo(mail, password);
        Console.WriteLine("Loaded User Info");
        return info;
    }
    public static Dictionary<string, string> GetFolders()
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        using var stream = File.Open(@"../../../data.xlsx", FileMode.Open, FileAccess.Read);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
        {
            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
            {
                UseHeaderRow = true
            }
        });
        var table = result.Tables[1];
        var dictionary = new Dictionary<string, string>();
        foreach (DataRow row in table.Rows)
        {
            string? key = row[0].ToString();
            string? value = row[1].ToString();
            if (key != null || value != null)
            {
                List<string> wordsList = value.Split(' ').ToList();
                foreach (var word in wordsList)
                {
                    dictionary.Add(word.ToUpper(), key);
                }
            }

        }
        stream.Close();
        Console.WriteLine("Loaded Folders and Keys");
        return dictionary;
    }

    public static ImapClient ConnectToMail()
    {
        var client = new ImapClient();
        try
        {
            client.Connect("imap.mail.ru", 993, true);
            Console.WriteLine("Client connection: " + client.IsConnected);
            return client;
        }
        catch (Exception)
        {
            Console.WriteLine("Error: Invalid server");
            Environment.Exit(-1);
            throw; //try to fix
        }
    }

    public static ImapClient AuthentificateUser(ImapClient client, UserInfo info)
    {
        try
        {
            client.Authenticate(info.mail, info.password);
            Console.WriteLine("Client Authentification: " + client.IsAuthenticated);
            return client;
        }
        catch (Exception)
        {
            Console.WriteLine("Error: invalid mail or password");
            Environment.Exit(-1);
            throw; //try to fix
        }
    }

    public static void CheckFolders(ImapClient client, Dictionary<string, string> dictionary)
    {
        HashSet<string> folders = new HashSet<string>(dictionary.Values);
        var inbox = client.GetFolder("INBOX");
        foreach (var folder in folders)
        {
            try
            {
                if (inbox.GetSubfolder(folder) != null)
                {
                    Console.WriteLine("Folder " + folder + " exists!");
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Cant find folder " + folder);
                var newFolder = client.GetFolder("INBOX");
                newFolder.Create(folder, false);
                Console.WriteLine("Created folder " + folder);
            }   
        }

        try
        {
            if (inbox.GetSubfolder("Unsorted") != null)
            {
                Console.WriteLine("Folder Unsorted exists!");
            }
        }
        catch (Exception)
        {
            var newFolder = client.GetFolder("INBOX");
            newFolder.Create("Unsorted", false);
            Console.WriteLine("Created folder Unsorted");
        }
    }
    public static int MoveMessage(ImapClient client, UniqueId u, MimeMessage msg, Dictionary<string, string> dictionary)
    {
        var inbox = client.GetFolder("INBOX");
        var subject = msg.Subject.ToUpper();
        List<string> wordsList = subject.Split(' ').ToList();
        foreach (var word in wordsList)
        {
            if (dictionary.ContainsKey(word))
            {
                var destination = dictionary[word];
                inbox.MoveTo(u, inbox.GetSubfolder(destination));
                Console.WriteLine("Message [UID: " + u + "] MOVED TO " + destination);
                return 1;
            }
            else
            {
                inbox.MoveTo(u, inbox.GetSubfolder("Unsorted"));
                Console.WriteLine("Message [UID: " + u + "] MOVED TO Unsorted");
                return 0;
            }
        }

        return -1;
    }
}


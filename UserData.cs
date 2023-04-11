using ExcelDataReader;
using System.Data;

namespace mailFolderParser
{
    internal class UserData
    {
        private readonly string? Mail;

        private readonly string? Password;

        private Dictionary<string, string> Folders;


        public UserData()
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

            this.Mail = mail;
            this.Password = password;

            table = result.Tables[1];
            //Making dictionary
            var dictionary = new Dictionary<string, string>();
            foreach (DataRow row in table.Rows)
            {
                string key = row[0].ToString();
                string value = row[1].ToString();
                if (key != null || value != null)
                {
                    List<string> wordsList = value.Split(' ').ToList();
                    foreach (var word in wordsList)
                    {
                        dictionary.Add(word.ToUpper(), key);
                    }
                }

            }
            this.Folders = dictionary;
            stream.Close();
        }

        public string GetMail() => this.Mail;
        public string GetPassword() => this.Password;
        public Dictionary<string, string> GetFolders() => this.Folders;
    }
}

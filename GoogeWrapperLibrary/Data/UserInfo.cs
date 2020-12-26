using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;

namespace GoogeWrapperLibrary.Data
{
    public static class UserInfo
    {
        public const string  outPath= "UserInfo.json";
        public const string credentials = "credentials.json";
        static UserInfo()
        {
           
            using (var readStream=new FileStream(credentials, FileMode.Open, FileAccess.Read))
            {
                User = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(readStream).Secrets,
                    new string[] { SheetsService.Scope.Spreadsheets },
                    "user",
                    CancellationToken.None,
                    new FileDataStore(outPath)).Result;  
            }
        }
        public static UserCredential User { get; set; }
    }
}

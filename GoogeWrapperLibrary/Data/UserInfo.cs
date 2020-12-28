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
        public const string credentials = "Wrapper-b6ea4c6d71a5.json";
        public static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static UserInfo()
        {
           
            using (var readStream=new FileStream(credentials, FileMode.Open, FileAccess.Read))
            {
                User = GoogleCredential.FromStream(readStream).CreateScoped(Scopes);
            }
        }
        public static GoogleCredential User { get; set; }
    }
}

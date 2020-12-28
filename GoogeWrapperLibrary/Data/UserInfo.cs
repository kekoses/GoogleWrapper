using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;

namespace GoogeWrapperLibrary.Data
{
    public class UserInfo
    {
        public static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        public GoogleCredential User { get; set; }
    }
}

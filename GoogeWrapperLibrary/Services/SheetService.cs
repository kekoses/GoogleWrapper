using GoogeWrapperLibrary.Data;
using GoogeWrapperLibrary.Data.Models;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogeWrapperLibrary.Services
{
    public class SheetService
    {
        private UserCredential _currentUser;
        public SheetService(string settingPath)
        {
            _currentUser = UserInfo.User;
        }
      
    }
}

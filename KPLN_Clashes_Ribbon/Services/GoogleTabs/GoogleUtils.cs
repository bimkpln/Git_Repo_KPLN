using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using static Google.Apis.Sheets.v4.SheetsService;

namespace KPLN_Clashes_Ribbon.Services.GoogleTabs
{
    internal class GoogleUtils
    {
        private static readonly string[] _scopes = { SheetsService.Scope.Spreadsheets };
        private static readonly string _applicationName = "KPLN_Google_App";
        
        private static readonly string _kplnLoaderSpreadSheetId = "1sFx8Vd_n9RI9rNFUjtiJcfGK1rJo8v4Bb993dnrub-I";
        private static readonly string _kplnLoaderStartRange = "Users!A2:K";

        private static GoogleCredential _loaderCredential;
        private static SheetsService _loaderSheetsService;

        public static GoogleCredential LoaderCredential 
        {
            get
            {
                if (_loaderCredential == null)
                {
                    using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                    {
                        _loaderCredential = GoogleCredential.FromStream(stream).CreateScoped(_scopes);
                    }
                }

                return _loaderCredential;
            }
        }

        public static SheetsService LoaderSheetsService
        {
            get
            {
                if (_loaderSheetsService == null)
                {
                    _loaderSheetsService = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = LoaderCredential,
                        ApplicationName = _applicationName
                    });
                }

                return _loaderSheetsService;
            }
        }
    }
}

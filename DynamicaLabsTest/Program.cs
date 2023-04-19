
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace GoogleDriveManager
{
    class Program
    {
       
        static void Main(string[] args)
        {
            GoogleDriveManager.WriteDocsListToSpreadsheet();

        }

    }
}

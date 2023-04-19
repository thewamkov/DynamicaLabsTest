
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace GoogleDriveManager
{
    class Program
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets, DriveService.Scope.DriveReadonly };
        static string ApplicationName = "Google Drive Sync";
        static string SpreadsheetName = "My Spreadsheet";
        static string SheetName = "Sheet1";
        static string CredentialsPath = "credentials.json";
        static string FolderId = "folderId"; // Replace with your Google Drive folder ID

        static void Main(string[] args)
        {
            GoogleDriveManager.WriteDocsListToSpreadsheet();

        }

    }
}
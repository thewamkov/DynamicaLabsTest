using System;
using Google.Apis.Auth.OAuth2;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Threading;
using Google.Apis.Util.Store;
using RestSharp;
using Google.Apis.Drive.v3;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using Google.Apis.Drive.v3.Data;
using File = Google.Apis.Drive.v3.Data.File;
using Google.Apis.Services;

namespace GoogleDriveManager
{
    public static class GoogleDriveManager
    {
        static private Configuration _config;
        static private string _Token;
        static private string _TokenType;

        static private File[] files;
        static private List<object> names;

        static private SheetsService sheetsService;
        static private DriveService driveService;
        static private UserCredential credential;

        private static Timer updateTimer;

        // Static constructor
        static GoogleDriveManager()
        {
            //Initialize config
            _config = ConfigurationManager.OpenExeConfiguration(Assembly.GetExecutingAssembly().Location);


            // Login 
            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                new ClientSecrets
                {
                    ClientId = _config.AppSettings.Settings["client_id"].Value,
                    ClientSecret = _config.AppSettings.Settings["client_secret"].Value
                },
                new[] { DriveService.Scope.Drive },
                "user",
                CancellationToken.None,
                new FileDataStore(@"C:\Studying", true)).Result;

            _Token = credential.Token.AccessToken;
            _TokenType = credential.Token.TokenType;


            sheetsService = InitializeSheetsService();

            driveService = InitializeDriveService();



        }

        public static void WriteDocsListToSpreadsheet()
        {
            // Start the update timer to run the code every 1 minute
            updateTimer = new Timer(UpdateTimerCallback, null, TimeSpan.Zero, TimeSpan.FromMinutes(15));


            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        private static void UpdateTimerCallback(object state)
        {
            Console.WriteLine("Updating spreadsheet...");

            try
            {
                // Run CreateOrUpdateSpreadsheet on a separate thread
                ThreadPool.QueueUserWorkItem(state =>
                {
                    CreateOrUpdateSpreadsheet(_config.AppSettings.Settings["document_name"].Value);
                });

                Console.WriteLine("Spreadsheet update started.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating spreadsheet: {ex.Message}");
            }
        }

        private static void CreateOrUpdateSpreadsheet(string spreadsheetName)
        {

            List<File> allFiles = RetrieveAllFiles();
            var spreadsheet = allFiles.FirstOrDefault(el => el.Name == spreadsheetName);
            var spreadsheetId = spreadsheet == null ? null : spreadsheet.Id;


            if (String.IsNullOrEmpty(spreadsheetId))
                spreadsheetId = CreateSpreadsheet(spreadsheetName);
            UpdateSpreadsheet(spreadsheetId, allFiles);



        }


        static void UpdateSpreadsheet(string spreadsheetId, List<File> allFiles)
        {
            try
            {
                var data = new List<IList<object>>() { new List<object>() { "Name", "Id" } };
                allFiles.ForEach(file => data.Add(new List<object>() { file.Name, file.Id }));


                // Clear existing data in the sheet
                var clearRequest = new BatchClearValuesRequest();
                sheetsService.Spreadsheets.Values.BatchClear(clearRequest, spreadsheetId).Execute();

                // Update data in the sheet
                var body = new ValueRange { Values = data };
                var updateRequest = sheetsService.Spreadsheets.Values.Update(body, spreadsheetId, "A1:C");
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var updateResponse = updateRequest.Execute();

                Console.WriteLine("Spreadsheet updated at: " + DateTime.Now);
            }
            catch (Google.GoogleApiException ex)
            {

                Console.WriteLine("An error occurred while updating the spreadsheet: " + ex.Message);

            }
        }



        public static string CreateSpreadsheet(string DocumentName)
        {

            // Create a new spreadsheet
            var spreadsheet = new Spreadsheet()
            {
                Properties = new SpreadsheetProperties()
                {
                    Title = DocumentName
                }
            };

            try
            {
                var createRequest = sheetsService.Spreadsheets.Create(spreadsheet);
                var createResponse = createRequest.Execute();

                return createResponse.SpreadsheetId;
            }
            catch (Exception ex)
            {

                Console.WriteLine("An error occurred while creating the spreadsheet: " + ex.Message);
                return null;
            }
        }


        public static List<File> RetrieveAllFiles()
        {
            List<File> result = new List<File>();

            FilesResource.ListRequest request = driveService.Files.List();

            do
            {
                try
                {
                    FileList files = request.Execute();


                    result.AddRange(files.Files);
                    request.PageToken = files.NextPageToken;
                }
                catch (Exception e)
                {
                    Console.WriteLine("An error occurred: " + e.Message);
                    request.PageToken = null;
                }
            } while (!String.IsNullOrEmpty(request.PageToken));
            return result;
        }

        public static SheetsService InitializeSheetsService()
        {


            return new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential
            }); ;
        }

        public static DriveService InitializeDriveService()
        {


            return new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential
            }); ;
        }


    }
}
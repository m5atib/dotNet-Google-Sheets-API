using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace dotNet_Google_Sheets_API
{
    public class GoogleSheetsAPI
    {
    
        //here put your sheet and client informtion.
        private const string SpreadsheetId = "";
        private const string clientId = "";
        private const string clientSecret = "";
        private const string APPname = "";

        private SpreadsheetsResource.ValuesResource serviceValues;
        
       
     

        //this method to connect the app with your google api account.
        private SheetsService GetSheetsService()
        {

            try
            {
                string[] scopes = { SheetsService.Scope.Spreadsheets };
                var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    new ClientSecrets { ClientId = clientId, ClientSecret = clientSecret },
                    scopes, Environment.UserName, CancellationToken.None, new FileDataStore("MyAppsToken")).Result;

                var serviceInitializer = new BaseClientService.Initializer()
                { HttpClientInitializer = credential, ApplicationName = APPname,};

                return new SheetsService(serviceInitializer);
            }
            catch
            {
                Console.WriteLine("Error with Getting Sheets Service.");
                return null;
            }
            
        }
        
        //to get data from sheet pages
        public IList<IList<object>> GetValuesInSheet(string SheetName)
        {
            try
            {
                serviceValues = GetSheetsService().Spreadsheets.Values;
                var response = serviceValues.Get(SpreadsheetId, SheetName + "!A:Z").Execute();
                var values = response.Values;

                if (values == null || !values.Any())
                {

                    return null;
                }

                return values;
            }
            catch
            {
                Console.WriteLine("Error with getting values from sheet or your name sheet page is worng.");
                return null;
            }

        }

        //to write data on sheet page
        public void WriteAsync(string WriteRange, List<object> row)
        {
            serviceValues = GetSheetsService().Spreadsheets.Values;
            var valueRange = new ValueRange { Values = new List<IList<object>> { row } };
            var update = serviceValues.Update(valueRange, SpreadsheetId, WriteRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
           var response =  update.Execute();

        }


    }

}
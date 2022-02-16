 public static class GoogleDataStudioClient
    {
        public static SheetsService GetGoogleSpreadSheetService(IUoW uow, int clientId, string sheetId, int apiKeySecretId, int accountTypeId)
        {
            try
            {
                var clientAccount = uow.ClientAccountRepository
                                    .Find(f => f.ClientId == clientId && f.AccountTypeId == accountTypeId && f.IsActive == true)
                                    .FirstOrDefault();

                if (clientAccount != null)
                {
                    var clientAPISecret = uow.ApiKeySecretRepository.Find(x => x.Id == apiKeySecretId).FirstOrDefault();
                    if (clientAPISecret != null)
                    {

                        var model = new GoogleUserCredential()
                        {
                            AccessToken = clientAccount.AccessToken,
                            RefreshToken = clientAccount.RefreshToken,
                            ClientId = clientAPISecret.ClientId,
                            ClientSecret = clientAPISecret.ClientSecret,
                            ApplicationName = "Flightdeck.Digital Service"
                        };
                        var serviceInitializer = Credential.GetGoogleServiceInitializer(model);
                        var service = new SheetsService(serviceInitializer);
                        return service;
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLogException(ex, FDGoogleMyBusinessLogType.FDGoogleDataStudioService);
            }
            return null;

        }
        public static GoogleDataStudioResponse GetGoogleSpreadSheetDetail(IUoW uow, int clientId, string sheetId, int apiKeySecretId = 20, int accountTypeId = 27)
        {
            var response = new GoogleDataStudioResponse();
            try
            {
                var service = GetGoogleSpreadSheetService(uow, clientId, sheetId, apiKeySecretId, accountTypeId);
                var spreadsheet = service.Spreadsheets.Get(sheetId).Execute();

                response.Spreadsheet = spreadsheet;
                response.SheetFileName = spreadsheet.Properties.Title;
                response.SheetId = spreadsheet.SpreadsheetId;
                response.HasError = false;
            }
            catch (Exception ex)
            {
                response.HasError = true;
                WriteLogException(ex, FDGoogleMyBusinessLogType.FDGoogleDataStudioService);
            }
            return response;

        }
        public static void AddSheetAtParticularIndex(string sheetName, int index, IUoW uow, int clientId, string sheetId, int apiKeySecretId = 20, int accountTypeId = 27)
        {
            try
            {
                var service = GetGoogleSpreadSheetService(uow, clientId, sheetId, apiKeySecretId, accountTypeId);
                // Add new Sheet
                var addSheetRequest = new AddSheetRequest();
                addSheetRequest.Properties = new SheetProperties();
                addSheetRequest.Properties.Title = sheetName;
                addSheetRequest.Properties.Index = index;

                BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
                batchUpdateSpreadsheetRequest.Requests = new List<Request>();
                batchUpdateSpreadsheetRequest.Requests.Add(new Request
                {
                    AddSheet = addSheetRequest
                });

                var batchUpdateRequest = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, sheetId);
                var response = batchUpdateRequest.Execute();

            }
            catch (Exception ex)
            {


                throw;
            }

        }
        public static void UpdateSheetAtParticularIndex(string newSheetName, string sheetName, int index, IUoW uow, int clientId, string sheetId, int apiKeySecretId = 20, int accountTypeId = 27)
        {
            try
            {
                var service = GetGoogleSpreadSheetService(uow, clientId, sheetId, apiKeySecretId, accountTypeId);
                Spreadsheet spreadsheet = service.Spreadsheets.Get(sheetId).Execute();
                Sheet sheet = spreadsheet.Sheets.Where(s => s.Properties.Title == sheetName).FirstOrDefault();
                int curretSheetId = (int)sheet.Properties.SheetId;
                var dd = sheet.Data;
                //getting all data from sheet
                var valTest = service.Spreadsheets.Values.Get(sheetId, $"{sheetName}!A2:D2").Execute();
                var valvalues = valTest.Values;
                //adding to specific row
                InsertRangeRequest insRow = new InsertRangeRequest();
                insRow.Range = new GridRange()
                {
                    SheetId = curretSheetId,
                    StartRowIndex = 3,
                    EndRowIndex = 4,



                };
                insRow.ShiftDimension = "ROWS";
                //sorting
                SortRangeRequest sorting = new SortRangeRequest();
                sorting.Range = new GridRange()
                {
                    SheetId = curretSheetId,
                    StartColumnIndex = 0,   // sorted by firstcolumn for all data after 1st row                 
                    StartRowIndex = 1
                };
                sorting.SortSpecs = new List<SortSpec>
                {
                    new SortSpec
                    {
                        SortOrder = "DESCENDING"
                    }
                };
                //
                InsertDimensionRequest insertRow = new InsertDimensionRequest();
                insertRow.Range = new DimensionRange()
                {
                    SheetId = curretSheetId,
                    Dimension = "ROWS",
                    StartIndex = 1,// 0 based
                    EndIndex = 2
                };
                var oblistt = new List<object>() { "Helloins", "This4ins", "was4ins", "insertd4ins" };
                PasteDataRequest data = new PasteDataRequest
                {
                    Data = string.Join(",", oblistt),
                    Delimiter = ",",
                    // data gets inserted form this cordinate point( i.e column and row) index and to the right
                    Coordinate = new GridCoordinate
                    {
                        ColumnIndex = 0, // 0 based Col A is 0, col B is 1 and so on
                        RowIndex = 2, //
                        SheetId = curretSheetId
                    },
                };

                BatchUpdateSpreadsheetRequest r = new BatchUpdateSpreadsheetRequest()
                {
                    Requests = new List<Request>
                {
                    //new Request{ InsertDimension = insertRow },
                    //new Request{ PasteData = data },
                    //new Request { InsertRange = insRow},
                    new Request { SortRange = sorting }
                }
                };

                BatchUpdateSpreadsheetResponse response1 = service.Spreadsheets.BatchUpdate(r, sheetId).Execute();
                //adding data to sheet
                var valueRange = new ValueRange();

                var oblist = new List<object>() { "Hello466", "This466", "was466", "insertd466" };
                valueRange.Values = new List<IList<object>> { oblist };

                var appendRequest = service.Spreadsheets.Values.Append(valueRange, sheetId, $"{sheetName}!A:D");
                //appendRequest.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var appendReponse = appendRequest.Execute();
                //updating data to sheet
                var valueRangeUpdate = new ValueRange();

                var oblistUpdate = new List<object>() { "Hello4uqq", "This4uqq", "was4uqq", "insertd4uqq" };
                valueRangeUpdate.Values = new List<IList<object>> { oblistUpdate };

                var updateRequest = service.Spreadsheets.Values.Update(valueRangeUpdate, sheetId, $"{sheetName}!A10:D510");

                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var updateReponse = updateRequest.Execute();

                //deleting row from sheet
                var range = $"{sheetName}!A7:G7";
                var requestBody = new ClearValuesRequest();

                var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, sheetId, range);
                var deleteReponse = deleteRequest.Execute();


                // update new Sheet
                var updateSheetRequest = new UpdateSheetPropertiesRequest();
                updateSheetRequest.Properties = new SheetProperties();
                updateSheetRequest.Properties.Title = sheetName;//Sheet1
                updateSheetRequest.Properties.Index = index;
                updateSheetRequest.Properties.SheetId = curretSheetId;
                updateSheetRequest.Fields = "Index,Title";//fields that needs to be updated.* means all fields from that sheet

                BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
                batchUpdateSpreadsheetRequest.Requests = new List<Request>();
                batchUpdateSpreadsheetRequest.Requests.Add(new Request
                {
                    UpdateSheetProperties = updateSheetRequest
                });

                var batchUpdateRequest = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, sheetId);
                var response = batchUpdateRequest.Execute();

            }
            catch (Exception ex)
            {


                throw;
            }

        }
        public static void ReadGoogleSheet(string newSheetName, string sheetName, int index, IUoW uow, int clientId, string sheetId, int apiKeySecretId = 20, int accountTypeId = 27)
        {
            try
            {
                var service = GetGoogleSpreadSheetService(uow, clientId, sheetId, apiKeySecretId, accountTypeId);
                Spreadsheet spreadsheet = service.Spreadsheets.Get(sheetId).Execute();
                Sheet sheet = spreadsheet.Sheets.Where(s => s.Properties.Title == sheetName).FirstOrDefault();
                int curretSheetId = (int)sheet.Properties.SheetId;
                var dd = sheet.Data;
                // update new Sheet
                var updateSheetRequest = new UpdateSheetPropertiesRequest();
                updateSheetRequest.Properties = new SheetProperties();
                updateSheetRequest.Properties.Title = sheetName;//Sheet1
                updateSheetRequest.Properties.Index = index;
                updateSheetRequest.Properties.SheetId = curretSheetId;
                updateSheetRequest.Fields = "Index,Title";//fields that needs to be updated.* means all fields from that sheet

                BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
                batchUpdateSpreadsheetRequest.Requests = new List<Request>();
                batchUpdateSpreadsheetRequest.Requests.Add(new Request
                {
                    UpdateSheetProperties = updateSheetRequest
                });

                var batchUpdateRequest = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, sheetId);
                var response = batchUpdateRequest.Execute();

            }
            catch (Exception ex)
            {


                throw;
            }

        }
        //public static GoogleDataStudioResponse GetGoogleSpreadSheetDetail(IUoW uow, int clientId, string sheetId, int apiKeySecretId = 20, int accountTypeId = 27)
        //{
        //    var response = new GoogleDataStudioResponse();
        //    try
        //    {
        //        var clientAccount = uow.ClientAccountRepository
        //                            .Find(f => f.ClientId == clientId && f.AccountTypeId == accountTypeId && f.IsActive == true)
        //                            .FirstOrDefault();

        //        if (clientAccount != null)
        //        {
        //            var clientAPISecret = uow.ApiKeySecretRepository.Find(x => x.Id == apiKeySecretId).FirstOrDefault();
        //            if (clientAPISecret != null)
        //            {

        //                var model = new GoogleUserCredential()
        //                {
        //                    AccessToken = clientAccount.AccessToken,
        //                    RefreshToken = clientAccount.RefreshToken,
        //                    ClientId = clientAPISecret.ClientId,
        //                    ClientSecret = clientAPISecret.ClientSecret,
        //                    ApplicationName = "Flightdeck.Digital Service"
        //                };
        //                var serviceInitializer = Credential.GetGoogleServiceInitializer(model);
        //                var service = new SheetsService(serviceInitializer);
        //                var spreadsheet = service.Spreadsheets.Get(sheetId).Execute();

        //                response.Spreadsheet = spreadsheet;
        //                response.SheetFileName = spreadsheet.Properties.Title;
        //                response.SheetId = spreadsheet.SpreadsheetId;
        //                response.HasError = false;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        response.HasError = true;
        //        WriteLogException(ex, FDGoogleMyBusinessLogType.FDGoogleDataStudioService);
        //    }
        //    return response;

        //}
    }
    public class GoogleDataSheet
    {
        public SheetsService _service;
        public GoogleDataSheet()
        {

        }
        public GoogleDataSheet(IUoW uow, int clientId, string sheetId, int apiKeySecretId, int accountTypeId)
        {
            _service = this.GetGoogleSpreadSheetService(uow, clientId, sheetId, apiKeySecretId, accountTypeId);
        }
        public SheetsService GetGoogleSpreadSheetService(IUoW uow, int clientId, string sheetId, int apiKeySecretId, int accountTypeId)
        {
            try
            {
                var clientAccount = uow.ClientAccountRepository
                                    .Find(f => f.ClientId == clientId && f.AccountTypeId == accountTypeId && f.IsActive == true)
                                    .FirstOrDefault();

                if (clientAccount != null)
                {
                    var clientAPISecret = uow.ApiKeySecretRepository.Find(x => x.Id == apiKeySecretId).FirstOrDefault();
                    if (clientAPISecret != null)
                    {

                        var model = new GoogleUserCredential()
                        {
                            AccessToken = clientAccount.AccessToken,
                            RefreshToken = clientAccount.RefreshToken,
                            ClientId = clientAPISecret.ClientId,
                            ClientSecret = clientAPISecret.ClientSecret,
                            ApplicationName = "Flightdeck.Digital Service"
                        };
                        var serviceInitializer = Credential.GetGoogleServiceInitializer(model);
                        var service = new SheetsService(serviceInitializer);
                        return service;
                    }
                }
            }
            catch (Exception ex)
            {
                GSheetReportingUtility.WriteLog($"GoogleSheet Service Error. Error Message is  {ex.Message}", GSheetReportingLogType.GSheetReportingService);
                //WriteLogException(ex, FDGoogleMyBusinessLogType.FDGoogleDataStudioService);
            }
            return null;

        }
        public (bool justCreated, bool hasError, string message) CheckOrCreateSheet(string sheetName, int sheetIndex, string sheetId)
        {

            try
            {
                //Get spreadsheet and check if sheet exist or not
                Spreadsheet spreadsheet = _service.Spreadsheets.Get(sheetId).Execute();
                Sheet sheet = spreadsheet.Sheets.Where(s => s.Properties.Title == sheetName).FirstOrDefault();

                // if sheet is not available create new sheet at specified index
                if (sheet == null)
                {
                    var addSheetRequest = new AddSheetRequest();
                    addSheetRequest.Properties = new SheetProperties();
                    addSheetRequest.Properties.Title = sheetName;
                    addSheetRequest.Properties.Index = sheetIndex;

                    BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
                    batchUpdateSpreadsheetRequest.Requests = new List<Request>();
                    batchUpdateSpreadsheetRequest.Requests.Add(new Request
                    {
                        AddSheet = addSheetRequest
                    });

                    var batchUpdateRequest = _service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, sheetId);
                    var response = batchUpdateRequest.Execute();
                    return (justCreated: true, hasError: false, message: "New sheet created successfully");
                }

                else
                {
                    // checking if it has headers or not assuming first row is header
                    var header = _service.Spreadsheets.Values.Get(sheetId, $"{sheetName}!1:1").Execute();
                    var valvalues = header.Values;
                    if (valvalues == null || valvalues.Count == 0)
                        return (justCreated: true, hasError: false, message: "Existing sheet without header");
                }


                return (justCreated: false, hasError: false, message: "");

            }
            catch (Exception ex)
            {
                return (justCreated: true, hasError: true, message: ex.Message);

            }

        }
        public (bool hasError, string message) InsertRowAtSpecifiedPlace(string sheetName, string sheetId, int startRowIndex, int endRowIndex)
        {
            try
            {
                Spreadsheet spreadsheet = _service.Spreadsheets.Get(sheetId).Execute();
                Sheet sheet = spreadsheet.Sheets.Where(s => s.Properties.Title == sheetName).FirstOrDefault();
                int curretSheetId = (int)sheet.Properties.SheetId;

                InsertRangeRequest insRow = new InsertRangeRequest();
                insRow.Range = new GridRange()
                {
                    SheetId = curretSheetId,
                    StartRowIndex = startRowIndex,
                    EndRowIndex = endRowIndex,
                };
                insRow.ShiftDimension = "ROWS";
                BatchUpdateSpreadsheetRequest request = new BatchUpdateSpreadsheetRequest()
                {
                    Requests = new List<Request>
                    {
                        new Request { InsertRange = insRow}
                    }
                };
                BatchUpdateSpreadsheetResponse response1 = _service.Spreadsheets.BatchUpdate(request, sheetId).Execute();
                return (hasError: false, message: "Row inserted successfully");
            }
            catch (Exception ex)
            {
                return (hasError: true, message: ex.Message);

            }
        }
        public (bool hasError, string message) AddValuesToSheet(string sheetName, string sheetId, Dictionary<string, string> dict)
        {

            try
            {
                var lastColumn = ColumnIndexToColumnLetter(dict.Count);
                var range = $"{sheetName}!A:{lastColumn}";
                var objectList = new List<object>();

                //using Sheet Header as key and getting its corresponding data from dictionary value
                var header = _service.Spreadsheets.Values.Get(sheetId, $"{sheetName}!A1:{lastColumn}1").Execute();
                var valvalues = header.Values;
                var keys = new List<string>();
                foreach (var item in valvalues)
                {
                    keys = item.Select(s => (string)s).ToList();
                }
                foreach (var item in keys)
                {
                    var value = dict[item];
                    objectList.Add(value);
                }

                //Inserting values directly without worrying about orders
                //var sheetValue = dict.Values.ToList();
                //foreach (var item in sheetValue)
                //{   
                //    objectList.Add(item);
                //}

                var valueRange = new ValueRange();
                valueRange.Values = new List<IList<object>> { objectList };

                //var oblist = new List<object>() { "Hello", "Sudip", "Rana", "Bhat" };
                //valueRange.Values = new List<IList<object>> { oblist };

                var appendRequest = _service.Spreadsheets.Values.Append(valueRange, sheetId, range);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var appendReponse = appendRequest.Execute();
                return (hasError: false, message: "Data Added successfully");

            }
            catch (Exception ex)
            {
                return (hasError: true, message: ex.Message);
            }

        }
        public (bool hasError, string message) UpdateValuesToSheet(int row, string sheetName, string sheetId, Dictionary<string, string> dict)
        {

            try
            {
                var lastColumn = ColumnIndexToColumnLetter(dict.Count);
                var range = $"{sheetName}!A{row}:{lastColumn}{row}";
                var objectList = new List<object>();

                //using Sheet Header as key and getting its corresponding data from dictionary value
                var header = _service.Spreadsheets.Values.Get(sheetId, $"{sheetName}!A1:{lastColumn}1").Execute();
                var valvalues = header.Values;
                var keys = new List<string>();
                foreach (var item in valvalues)
                {
                    keys = item.Select(s => (string)s).ToList();
                }
                foreach (var item in keys)
                {
                    var value = dict[item];
                    objectList.Add(value);
                }
                var valueRange = new ValueRange();
                valueRange.Values = new List<IList<object>> { objectList };

                //var oblist = new List<object>() { "Hello", "Sudip", "Rana", "Bhat" };
                //valueRange.Values = new List<IList<object>> { oblist };

                var updateRequest = _service.Spreadsheets.Values.Update(valueRange, sheetId, range);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var updateReponse = updateRequest.Execute();
                return (hasError: false, message: "Data Updated successfully");

            }
            catch (Exception ex)
            {
                return (hasError: true, message: ex.Message);
            }

        }
        public (bool hasError, string message) DeleteValuesFromSheet(string sheetId, string sheetName, string startColumn, string endColumn, int? row = null)
        {

            try
            {
                var range = $"{sheetName}!{startColumn}:{endColumn}";
                if (row != null)
                    range = $"{sheetName}!{startColumn}{row}:{endColumn}{row}";

                var requestBody = new ClearValuesRequest();

                var deleteRequest = _service.Spreadsheets.Values.Clear(requestBody, sheetId, range);
                var deleteReponse = deleteRequest.Execute();
                return (hasError: false, message: "Data Deleted successfully");

            }
            catch (Exception ex)
            {
                return (hasError: true, message: ex.Message);
            }

        }
        public List<Dictionary<string, string>> GetAllValuesFromSheet(string sheetId, string sheetName, string startColumn, string endColumn, int? row = null)
        {
            var allData = new List<Dictionary<string, string>>();
            try
            {
                var range = $"{sheetName}!{startColumn}:{endColumn}";
                if (row != null)
                    range = $"{sheetName}!{startColumn}{row}:{endColumn}{row}";

                var objectList = new List<object>();

                var sheetData = _service.Spreadsheets.Values.Get(sheetId, range).Execute();
                var valvalues = sheetData.Values;
                var keys = new List<string>();

                int count = 1;
                foreach (var item in valvalues)
                {
                    if (count == 1)
                    {
                        keys = item.Select(s => (string)s).ToList();
                    }
                    if (count > 1)
                    {
                        var values = new List<string>();
                        values = item.Select(s => (string)s).ToList();
                        var objResult = new Dictionary<string, string>();
                        for (int j = 0; j < values.Count; j++)
                            objResult.Add(keys[j], values[j]);

                        allData.Add(objResult);
                    }
                    count++;

                }
                // To get all data from list of dictionary that has same key (To get all data from specific column)
                //var singleColumnData = allData.Select(dict => dict["KeyName"]).Distinct().ToList();



            }
            catch (Exception ex)
            {

            }
            return allData;

        }
        public static string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }
    }

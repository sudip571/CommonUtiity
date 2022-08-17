public static string CreateCSVFile(IUoW uow, int clientId, long? DomainId, DateTime startDate, DateTime endDate,
	        int? SearchEngineId, string CountryCode, int frequency, bool isMobie)
        {

	        try
	        {
		        Logger("CreateCSVFile : ", LogType.RankingSnapshotEmailSchedulerService);

		        var p = new object[]
		        {
			        new SqlParameter("@ClientId", SqlDbType.Int) {Value = clientId},
			        new SqlParameter("@DomainId", SqlDbType.Int) {Value = DomainId},
			        new SqlParameter("@Frequency", SqlDbType.Int) {Value = frequency},
			        new SqlParameter("@FromDate", SqlDbType.DateTime) {Value = startDate},
			        new SqlParameter("@ToDate", SqlDbType.DateTime) {Value = endDate},
			        new SqlParameter("@CountryCode", SqlDbType.VarChar) {Value = CountryCode},
			        new SqlParameter("@SearchEngineId", SqlDbType.Int) {Value = SearchEngineId},
			        new SqlParameter("@BucketId", SqlDbType.Int) {Value = 0}
		        };
		        //For Testing purpose
		        //var p = new object[]
		        //  {
		        //	 new SqlParameter("@ClientId", SqlDbType.Int) {Value = 2303},
		        //	 new SqlParameter("@DomainId", SqlDbType.Int) {Value = 10404265},
		        //	 new SqlParameter("@Frequency", SqlDbType.Int) {Value = 30},
		        //	 new SqlParameter("@FromDate", SqlDbType.DateTime) {Value = DateTime.Parse("2018-03-15")},
		        //	 new SqlParameter("@ToDate", SqlDbType.DateTime) {Value = DateTime.Parse("2018-05-21")},
		        //	 new SqlParameter("@CountryCode", SqlDbType.VarChar) {Value = "AU"},
		        //	 new SqlParameter("@SearchEngineId", SqlDbType.Int) {Value = 1},
		        //	 new SqlParameter("@BucketId", SqlDbType.Int) {Value = 0}
		        //  };

		        //Getting data from storedprocedure and binding those values dynamically on Dictionary
		        //var context = new MyDbContext()
		        //var cmd = context.Database.Connection.CreateCommand();
		        List<Dictionary<string, object>> items = new List<Dictionary<string, object>>();
		        //var uow = new UoW();
		        var context = uow.getContext();
		        context.Database.Connection.Open();
		        var cmd = context.Database.Connection.CreateCommand();
		        cmd.CommandText = @"[web].[GetRankingSnapshotExcelReportBySQLScriptForGoogleDataStudio]";
		        cmd.CommandType = CommandType.StoredProcedure;
		        cmd.Parameters.AddRange(p);

		        var reader = cmd.ExecuteReader();
		        //
		        while (reader.Read())
		        {
			        Dictionary<string, object> item = new Dictionary<string, object>();

			        for (int i = 0; i < reader.FieldCount; i++)
				        item[reader.GetName(i)] = (reader[i]);
			        items.Add(item);
		        }

		        Logger("items : " + items.Count, LogType.RankingSnapshotEmailSchedulerService);


		        if (items != null && items.Count > 0)
		        {
			        //items = items.Where(x => x.Keys[DomainId] == 4585);


			        //Create Excel File                    
			        bool needToAddHeader = true;
			        var j = 1;
			        using (var excelWorkBook = new XLWorkbook())
			        {
				        var workSheet =
					        excelWorkBook.Worksheets.Add($"RankingSnapShot-{startDate.ToString("yyyy-MM-dd")}");
				        for (int i = 0; i < items.Count; i++)
				        {
					        var dict = items[i];
					        var notToInclude = new List<string>() {"ClientId", "DomainId", "KeywordId"};
					        var keys = dict.Keys.ToList();
					        keys = keys.Except(notToInclude).ToList();
					        if (needToAddHeader)
					        {
						        var h = 1;
						        foreach (var key in keys)
						        {
							        var column = GetExcelColumnLetter(h);
							        workSheet.Cell($"{column}{j}").SetValue(key);
							        h++;

						        }

						        needToAddHeader = false;
						        j++;
						        i--;
					        }
					        else
					        {
						        var domId = dict["DomainId"].ToString();

						        if (domId == DomainId.ToString()) //DomainId.ToString())

						        {
							        var h = 1;
							        foreach (var key in keys)
							        {
								        var column = GetExcelColumnLetter(h);
								        workSheet.Cell($"{column}{j}").SetValue(dict[key]);
								        h++;
							        }

							        j++;
						        }
					        }

				        }

				        var fullPath =
					        $@"C:\\FlightdeckFile\\RankingSnapshotEmailScheduler\\Ranking-SnapShot-Report-{clientId}-{DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")}.csv";
				        //var fullPath = $@"C:\\FlightdeckLog\\EcommQueue\\MonthlyReport\\EcomReport-{clientId}-{DateTime.Now.ToString("yyyy-MM-dd")}.xlsx";//DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")
				        var fileInfo = new FileInfo(fullPath);
				        if (!fileInfo.Directory.Exists)
					        fileInfo.Directory.Create();
				        if (!File.Exists(fullPath))
				        {
					        File.Create(fullPath).Close();
				        }

				        //Converting worksheet to CSV format and saving it to given filepath
				        var lastCellAddress = workSheet.RangeUsed().LastCell().Address;
				        System.IO.File.WriteAllLines(fullPath, workSheet.Rows(1, lastCellAddress.RowNumber)
					        .Select(row => String.Join(",", row.Cells(1, lastCellAddress.ColumnNumber)
						        .Select(cell => cell.GetValue<string>()))
					        ));

				        // saving worksheet as EXCEL on given filepath
				        //excelWorkBook.SaveAs(fullPath);
				        Logger("fullPath : " + fullPath, LogType.RankingSnapshotEmailSchedulerService);
				        return fullPath;
			        }
		        }


	        }
	        catch (Exception ex)
	        {
		        Logger(ex, LogType.RankingSnapshotEmailSchedulerError);
		        return "";
	        }

	        return "";
        }

        public static string GetExcelColumnLetter(int colIndex)
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

public static SqlStatus BulkCopy(string destinationTable, DataTable table)
        {

            try
            {
                var connectionString = ConfigurationManager.ConnectionStrings["internaltoolsetEntities"].ConnectionString;
                connectionString = EFConnectionToADOConnection(connectionString);
                 var status = new SqlStatus();
                var timer = new Stopwatch();
                timer.Start();
                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            using (var bulkCopy = new SqlBulkCopy(connectionString))
                            {
                                //if (table.Rows.Count > 5000)
                                //{
                                //    bulkCopy.BatchSize = 5000;
                                //    bulkCopy.BulkCopyTimeout = 0;
                                //}
                                bulkCopy.BatchSize = 2000;
                                bulkCopy.BulkCopyTimeout = 0;
                                bulkCopy.DestinationTableName = destinationTable;
                                bulkCopy.WriteToServer(table);
                            }
                            transaction.Commit();
                            status.HasError = false;
                            status.Message = "Successfully inserted all data";
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            status.HasError = true;
                            status.Message = ex.Message;
                        }
                        timer.Stop();
                        status.TimeTaken = timer.Elapsed.TotalSeconds;
                        return status;
                    }
                }

            }
            catch (Exception ex)
            {

                throw;
            }
        }
        public static string EFConnectionToADOConnection(string connectionString)
        {
            //EF connectionString starts with 'metadata='
            if (connectionString.ToLower().StartsWith("metadata="))
            {
                EntityConnectionStringBuilder efBuilder = new EntityConnectionStringBuilder(connectionString);
                connectionString = efBuilder.ProviderConnectionString;
            }
            return connectionString;
        }

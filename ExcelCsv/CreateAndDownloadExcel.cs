 public byte[] PrepareContentForSearchTermDetailExport(List<SearchTermDetailModified> model)
        {
            try
            {
                int counter = 2;
                using (var excelWorkBook = new XLWorkbook())
                {
                    var workSheet = excelWorkBook.Worksheets.Add("Search Term Detail");
                    workSheet.Cell("A1").SetValue("Search Term");
                    workSheet.Cell("B1").SetValue("Paid Competitors");                    
                    workSheet.Cell("C1").SetValue("Estimated Organic Clicks");                    
                    workSheet.Cell("D1").SetValue("Average Organic Position");
                    workSheet.Cell("E1").SetValue("Top Paid Competitor");
                    workSheet.Cell("F1").SetValue("Min CPC ($)");
                    workSheet.Cell("G1").SetValue("Max CPC ($)");
                    workSheet.Cell("H1").SetValue("Organic CTR (%)");

                    //workSheet.Range("A1:F1").Style.Font.SetBold();

                    foreach (var item in model)
                    {
                        workSheet.Cell("A" + counter).SetValue(item.SearchTerm);
                        workSheet.Cell("B" + counter).SetValue(item.Competitors == -1 ? "-": item.Competitors.ToString());                     
                        workSheet.Cell("C" + counter).SetValue(item.EstimatedClicks == -1 ? "-" : item.EstimatedClicks.ToString());
                        workSheet.Cell("D" + counter).SetValue(item.AveragePosition == -1 ? "-" :item.AveragePosition.ToString());
                        workSheet.Cell("E" + counter).SetValue(item.TopCompetitor == "-1" ? "-": item.TopCompetitor);
                        workSheet.Cell("F" + counter).SetValue(item.MinCPC == -1 ? "-": item.MinCPC.ToString());
                        //workSheet.Cell("G" + counter).SetValue(item.MaxCPC == -1 ? "-":item.MaxCPC.ToString("N3"));
                        workSheet.Cell("G" + counter).SetValue(item.MaxCPC == -1 ? "-" : item.MaxCPC.ToString());
                        workSheet.Cell("H" + counter).SetValue(item.CTR == -1 ? "-" :(item.CTR * 100).ToString());
                        counter++;
                    }

                    //save Excel WorkBook to Memorystream
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        excelWorkBook.SaveAs(memoryStream);

                        //Convert MemoryStream to Byte array.
                        var bytes = memoryStream.ToArray();
                        return bytes;
                    }
                        // save excelworkbook to given filepath
                        // var fullPath = $@"C:\\FlightdeckLog\\EcommQueue\\MonthlyReport\\hahaazq-{DateTime.Now.ToString("yyyy-MM-dd")}.csv";
                        //var fullPath = $@"C:\\FlightdeckLog\\EcommQueue\\MonthlyReport\\EcomReport-{clientId}-{DateTime.Now.ToString("yyyy-MM-dd")}.xlsx";
                       // var fileInfo = new FileInfo(fullPath);
                       // if (!fileInfo.Directory.Exists)
                       //     fileInfo.Directory.Create();
                       // if (!File.Exists(fullPath))
                       // {
                       //     File.Create(fullPath).Close();
                       // }

                        //excelWorkBook.SaveAs(fullPath);

                        //Converting worksheet into CSV format and saving it to given path
                        //var lastCellAddress = workSheet.RangeUsed().LastCell().Address;
                        //System.IO.File.WriteAllLines(fullPath, workSheet.Rows(1, lastCellAddress.RowNumber)
                        //    .Select(row => String.Join(",", row.Cells(1, lastCellAddress.ColumnNumber)
                        //        .Select(cell => cell.GetValue<string>()))
                        //   ));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        
         //if (formatType == "csv")
 //           {
 //               extension = ".csv";
 //               contentType = "text/csv";
 //           }
 //           if (formatType == "excel")
 //           {
 //               extension = ".xlsx";
 //               contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
 //           }
 
  //var data = _fileContentPreparation.PrepareContentForSearchTermDetailExport(gridModel);
  //              var fileName = clientName + " - Search Term Detail" + DateTime.Now.Date.ToString("dd'-'MM'-'yyyy") + extension;
  //              return File(data, contentType, fileName);
        

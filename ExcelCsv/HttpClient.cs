//example to redirect to login consent
                   var scope = @"advertising::campaign_management";
                   var amazonAuthorizeAPI = "https://apac.account.amazon.com/ap/oa?scope={0}&response_type=code&client_id={1}&state=State&redirect_uri={2}";
                    amazonAuthorizeAPI = string.Format(amazonAuthorizeAPI, scope, apiClientId, redirectURL);
                    Response.Redirect(amazonAuthorizeAPI);
// after you get Code from Redirected url, do like following to get access token in CallBackAction method
                    var grantType = "authorization_code";
                    string data = "code={0}&client_id={1}&client_secret={2}&redirect_uri={3}&grant_type={4}";
                    data = string.Format(data, code, apiClientId, apiClientSecret, redirectUrl, grantType);
                    var responseObject = new
                    {
                        access_token = "",
                        refresh_token = "",
                        token_type = "",
                        expires_in = 0
                    };
                    var responseContent = GetAmazonAccessToken(data);
                    responseObject = JsonConvert.DeserializeAnonymousType(responseContent, responseObject);
//
public string GetAmazonAccessToken(string data)
        {
            try
            {
                #region requesting for access and refresh token
                var authURL = @"https://api.amazon.co.jp/auth/o2/token";
                var responseContent = string.Empty;
                _httpClient.DefaultRequestHeaders.Clear();
                _httpClient.DefaultRequestHeaders.Accept.Add(new
                System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/x-www-form-urlencoded"));
                var stringContent = new StringContent(data, Encoding.UTF8, "application/x-www-form-urlencoded");
                var response = _httpClient.PostAsync(authURL, stringContent).Result;
                if (response.IsSuccessStatusCode)
                {
                    responseContent = response.Content.ReadAsStringAsync().Result;
                }

                return responseContent;
                #endregion

            }
            catch (Exception)
            {

                throw;
            }
        }
        
   // example
                          IEnumerable<dynamic> sponsoredBrandResponse = null;
                         _httpClient.DefaultRequestHeaders.Clear();
                        _httpClient.DefaultRequestHeaders.Add("Amazon-Advertising-API-ClientId", clientAPISecret.ClientId);
                        _httpClient.DefaultRequestHeaders.Add("Amazon-Advertising-API-Scope", amazonProfile.ProfileId.ToString());
                        _httpClient.DefaultRequestHeaders.Add("Authorization", authToken);
                        _httpClient.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                        
                        //
                        var sbResponse = _httpClient.GetAsync(sbURL).Result;
						if (sbResponse.IsSuccessStatusCode)
						{
							var responseContent = sbResponse.Content.ReadAsStringAsync().Result;
							sponsoredBrandResponse = JsonConvert.DeserializeObject<IEnumerable<dynamic>>(responseContent);
							foreach (var sb in sponsoredBrandResponse)
							{
								var model = new AmazonCampaign()
								{
									CampaignId = sb.campaignId,
									CampaignName = sb.name,
									State = sb.state,
									CampaignType = "Sponsored Brand",
									StartDate = sb.startDate,
									EndDate = sb.endDate
								};
								amazonProfile.AmazonCampaigns.Add(model);
							}

						}
// for goolge  visit ths link   https://docs.informatica.com/integration-cloud/cloud-data-integration-connectors/current-version/google-sheets-connector/introduction-to-google-sheets-connector/administration-of-google-sheets-connector/generating-oauth-2-0-access-tokens.html

//other httpclient post example visit
// https://stackoverflow.com/questions/15176538/net-httpclient-how-to-post-string-value
// https://stackoverflow.com/questions/20005355/how-to-post-data-using-httpclient

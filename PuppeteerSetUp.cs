                    #region browser setup method1
                    // simple setup and 
                    //this is suitable if you are going to use it from web. By default it launches Chromium browser and everything needed for it will be downloaded to 
                    // folder .local-chromium inside project's Bin folder.
                    string[] argument = { "--no-sandbox" };
                    await new BrowserFetcher().DownloadAsync(BrowserFetcher.DefaultRevision);
                    var browser = await Puppeteer.LaunchAsync(new LaunchOptions
                    {
                        Headless = true,
                        IgnoreHTTPSErrors = true,
                        DefaultViewport = null,
                        Args = argument
                      
                    });
                    #endregion
                    
                     #region browser setup method2
                    //this is suitable if you are going to use it as windows service ( instead of web which has its own Bin and required dll files). By default it looks inside
                    // 'C:\Windows\system32\.local-chromium' for  required things to run but we might get UnAuthorized Access Error.
                    //So we can manually setup the path.At,first it takes longer time because it needs to unzip the downloaded zip file.
                    
                    string[] argument = { "--no-sandbox" };
                    var browserFetcher = new BrowserFetcher(new BrowserFetcherOptions
                    {
                        // this path can be setup anywhere
                        Path = "C:\\.local-chromium"
                    });                    
                    await browserFetcher.DownloadAsync(BrowserFetcher.DefaultRevision);                  
                    var browser = await Puppeteer.LaunchAsync(new LaunchOptions
                    {
                        Headless = false,
                        IgnoreHTTPSErrors = true,
                        DefaultViewport = null,
                        Args = argument,                        
                        ExecutablePath = browserFetcher.RevisionInfo(BrowserFetcher.DefaultRevision).ExecutablePath

                    });
                    #endregion
                    
                     #region browser setup method3
                    //this might be suitable than other two methods above.Instead of chromium we use chrome.
                    
                    string[] argument = { "--no-sandbox" };
                    //By default, it launches chromium browser but if you need chrome, just provide its executable path to LaunchOption.
                    var executablePath = @"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe";
                    var browser = await Puppeteer.LaunchAsync(new LaunchOptions
                    {
                        Headless = true,
                        IgnoreHTTPSErrors = true,
                        DefaultViewport = null,
                        Args = argument,
                        ExecutablePath = executablePath
                    });
                    #endregion
                    
                    //further set up
                     public  class BrowseRequest
    {
        public async Task<BrowserStatus> GetHtmlContentAsync(Browser browser,string url,string userAgent=null,int pageLoadWaiting=0,int pageRequestTimeOut=30000,bool loadImage=false,string waitingPath=null)
        {
            var status=new BrowserStatus();
            try
            {   
                if (string.IsNullOrEmpty(userAgent))
                    userAgent = Constant.COMMON_USER_AGENT;


                var page = await browser.NewPageAsync();
                page.DefaultTimeout = pageRequestTimeOut;
                await page.SetJavaScriptEnabledAsync(true);
                await page.SetUserAgentAsync(userAgent);
                //Disabling image helps webpage render faster but in some cases we need images
                if (!loadImage)
                {
                    await page.SetRequestInterceptionAsync(true);
                    page.Request += (sender, e) =>
                    {
                        if (e.Request.ResourceType == ResourceType.Image)
                            e.Request.AbortAsync();
                        else
                            e.Request.ContinueAsync();                        
                    };
                }
                // different WaitUntilNavigation value works for different sites.Test to find out which WaitUntilNavigation's value is suitable for given site.
                // sometimes you dont need WaitUntilNavigation  at all. wrong WaitUntilNavigation value causes redirect error,timeout error etc
                await page.GoToAsync(url,WaitUntilNavigation.DOMContentLoaded);
                //if (pageLoadWaiting > 0)
                //    await page.WaitForTimeoutAsync(pageLoadWaiting);
                if (!string.IsNullOrEmpty(waitingPath))
                    await page.WaitForSelectorAsync(waitingPath,new WaitForSelectorOptions {Timeout=10000 });
                var htmlDoc = await page.GetContentAsync();
                status.TargetUrl = page.Target.Url;
                await page.CloseAsync();
                status.status = true;
                status.StatusCode = 200;
                status.HtmlDocument = htmlDoc;
            }
            catch (Exception ex)
            {
                status.status = false;
                CheckErrorType(ex,status);               
                status.HtmlDocument = null;
                status.StatusMessage = ex.Message;

            }
            return status;
        }
        public void CheckErrorType(Exception ex,BrowserStatus status)
        {
            var exceptionMessage = ex.Message;
            //page is loaded but x-path could not be found
            if (exceptionMessage.StartsWith("waiting for selector"))
                status.StatusCode = 910;
            //page cannot be loaded
            else if (exceptionMessage.StartsWith("Timeout of"))
                status.StatusCode = 911;
            //other kind of error occured
            else
                status.StatusCode = 912;

            
        }
    }
}
                    
                    
                    
                    
                    

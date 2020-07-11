using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Logic.CS.BusinessLogic.Adthena
{
    public sealed class StringConstant
    {
        // Adthena API URLS
        public const string INFRINGEMENT_API = @"https://api.adthena.com/wizard/{0}/infringement/all";
        public const string SEARCH_TERM_DETAIL = @"https://api.adthena.com/wizard/{0}/search-term-detail/all";
        public const string SEARCH_TERM_OPPORTUNITIES = @"https://api.adthena.com/wizard/{0}/search-term-opportunities/all";

        // Session Varaiable Name
        public const string INFRINGEMENTGRID_SESSION = "InfringementGrid";
        public const string MISSINGORGANICGRID_SESSION = "MissingOrganicGrid";
        public const string SEARCH_TERM_DETAILGRID_SESSION = "SearchTermGrid";

    }
}
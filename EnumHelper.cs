using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;

namespace BCA_Web_Application.Areas.Adthena.Common
{
    public static class EnumHelper
    {
        public enum Status
        {
            [Description("success")]
            Success = 1,

            [Description("fail")]
            Fail = 2,

            [Description("unauthorized")]
            Unauthorized = 3,

            
        }
    }
    }
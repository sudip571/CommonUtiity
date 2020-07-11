using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Logic.CS.BusinessLogic.Adthena
{
    public static  class SessionManagement
    {
        public static T GetSession<T>(string key)
        {
            object sessionObject = HttpContext.Current.Session[key];
            if (sessionObject == null)
            {
                return default(T);
            }
            return (T)HttpContext.Current.Session[key];

        }

        public static void SaveOrUpdateSession<T>(string key, T entity)
        {
            object sessionObject = HttpContext.Current.Session[key];
            if (sessionObject != null)
            {
                DeleteSession(key);
            }            
            HttpContext.Current.Session[key] = entity;
        }

        public static void DeleteSession(string key)
        {
            HttpContext.Current.Session.Remove(key);
        }
    }
}
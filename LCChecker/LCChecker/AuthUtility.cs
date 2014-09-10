using LCChecker.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LCChecker
{
    public class AuthUtility
    {
        private static string _sessionKey = "user";

        public static User GetCurrentUser(HttpContextBase context)
        {
            var sessionValue = context.Session[_sessionKey];
            return sessionValue == null ? null : (User)sessionValue;
        }

        public static void SaveCurrentUser(HttpContextBase context, User user)
        {
            context.Session[_sessionKey] = user;
        }
    }
}
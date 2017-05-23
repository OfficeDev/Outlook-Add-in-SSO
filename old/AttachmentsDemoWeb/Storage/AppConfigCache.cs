// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using System.Collections.Generic;

namespace AttachmentsDemoWeb.Storage
{
    // This class implements a simplistic, in-memory cache
    // for user refresh tokens for demonstration purposes.
    // A production solution should cache refresh tokens in a 
    // more robust manner, such as in a secure database.
    public class AppConfigCache
    {
        // This dictionary maps state GUIDs to user email addresses
        private static Dictionary<string, string> authRequestGuidCache;
        // This dictionary maps user email addresses to AppConfig objects
        private static Dictionary<string, AppConfig> userConfigCache;

        private static void EnsureCache()
        {
            if (authRequestGuidCache == null)
                authRequestGuidCache = new Dictionary<string, string>();
            if (userConfigCache == null)
                userConfigCache = new Dictionary<string,AppConfig>();
        }
        public static AppConfig GetUserConfig(string userEmailAddress)
        {
            EnsureCache();
            AppConfig userConfig = null;

            if (userConfigCache.TryGetValue(userEmailAddress, out userConfig))
            {
                return userConfig;
            }

            return null;
        }

        public static void AddUserConfig(string userEmailAddress, AppConfig userConfig)
        {
            EnsureCache();
            // If the cache already contains an entry for this user, remove 
            // it.
            if (userConfigCache.ContainsKey(userEmailAddress))
            {
                userConfigCache.Remove(userEmailAddress);
            }
            userConfigCache.Add(userEmailAddress, userConfig);
        }

        public static void RemoveUserConfig(string userEmailAddress)
        {
            EnsureCache();
            userConfigCache.Remove(userEmailAddress);
        }

        public static string GetUserFromStateGuid(string stateGuid)
        {
            string userEmail = null;

            if (authRequestGuidCache.TryGetValue(stateGuid, out userEmail))
            {
                return userEmail;
            }

            return string.Empty;
        }

        public static void AddStateGuid(string stateGuid, string userEmail)
        {
            EnsureCache();
            // If the cache already contains an entry for this state GUID, 
            // remove it.
            if (authRequestGuidCache.ContainsKey(stateGuid))
            {
                authRequestGuidCache.Remove(stateGuid);
            }
            authRequestGuidCache.Add(stateGuid, userEmail);
        }

        public static void RemoveStateGuid(string stateGuid)
        {
            EnsureCache();
            authRequestGuidCache.Remove(stateGuid);
        }
    }
}

// MIT License: 

// Permission is hereby granted, free of charge, to any person obtaining 
// a copy of this software and associated documentation files (the 
// ""Software""), to deal in the Software without restriction, including 
// without limitation the rights to use, copy, modify, merge, publish, 
// distribute, sublicense, and/or sell copies of the Software, and to 
// permit persons to whom the Software is furnished to do so, subject to 
// the following conditions: 

// The above copyright notice and this permission notice shall be 
// included in all copies or substantial portions of the Software. 

// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, 
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. 
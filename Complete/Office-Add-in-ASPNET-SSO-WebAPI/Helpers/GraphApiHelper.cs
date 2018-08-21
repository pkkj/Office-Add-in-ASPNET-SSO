// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the root of the project.

/* 
    This file provides URLs to help get Microsoft Graph data. 
*/

using System;
using System.Globalization;
using Microsoft.Identity.Client;
using System.IdentityModel.Tokens;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Script.Serialization;
using System.Net.Http;
using System.Net.Http.Headers;
using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
using Office_Add_in_ASPNET_SSO_WebAPI.Models;
using System.Runtime.Serialization;
using System.IO;
using System.Text;
using System.Runtime.Serialization.Json;
namespace Office_Add_in_ASPNET_SSO_WebAPI.Helpers
{
    public class GraphTokenException : Exception
    {
        public string ErrorCode { get; }
        public string SubError { get; }
        public string Claims { get; }

        public override string ToString()
        {
            return base.ToString() + string.Format(CultureInfo.InvariantCulture, "\n\tErrorCode: {0}", ErrorCode);
        }

        public GraphTokenException(string errorCode, string errorMessage, string claims, string suberror)
            : base(errorMessage)
        {
            ErrorCode = errorCode;
            Claims = claims;
            SubError = suberror;
        }

    }

    [Serializable]
    [DataContract]
    public class GraphToken
    {
        [DataMember(Name = "result")]
        public string TokenType;

        [DataMember(Name = "scope")]
        public string Scope;

        [DataMember(Name = "expires_in")]
        public int ExpiresIn;

        [DataMember(Name = "access_token")]
        public string AccessToken;
    }

    /// <summary>
    /// Provides methods and strings for Microsoft Graph-specific endpoints.
    /// </summary>
    internal static class GraphApiHelper
    {
        // Microsoft Graph-related base URLs
        internal static string GetFilesUrl = @"https://graph.microsoft.com/v1.0/me/drive/root/children";

        internal static string GetMyInfoUrl = @"https://graph.microsoft.com/v1.0/me";

        internal static string GetOneDriveItemNamesUrl(string selectedProperties)
        {
            // Construct URL for the names of the folders and files.
            return GetFilesUrl + selectedProperties;
        }

        private const string tokenURLSegment = "/oauth2/v2.0/token";

        // Special implementation for acquiring the Graph token, because currently suberror is not supported by the MSAL library 
        internal static async Task<GraphToken> AcquireTokenOnBehalfOfAsync(string accessToken, string[] graphScopes)
        {
            using (var client = new HttpClient())
            {
                var values = new Dictionary<string, string>
                    {
                        { "client_id",  ConfigurationManager.AppSettings["ida:ClientID"] },
                        { "client_secret",  ConfigurationManager.AppSettings["ida:Password"] },
                        { "grant_type", "urn:ietf:params:oauth:grant-type:jwt-bearer" },
                        { "assertion", accessToken },
                        { "requested_token_use", "on_behalf_of" },
                        { "scope", string.Join(" ", graphScopes) }
                    };

                FormUrlEncodedContent content = new FormUrlEncodedContent(values);

                // Create and send the HTTP Request
                var request = new HttpRequestMessage(HttpMethod.Post, ConfigurationManager.AppSettings["ida:Tenant"] + tokenURLSegment);
                request.Content = content;

                request.Headers.Add("Accept", "application/json");

                using (HttpResponseMessage response = await client.SendAsync(request))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        string responseContent = await response.Content.ReadAsStringAsync();

                        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(GraphToken));
                        MemoryStream ms = new MemoryStream(Encoding.Unicode.GetBytes(responseContent));
                        GraphToken result = serializer.ReadObject(ms) as GraphToken;
                        return result;
                    }
                    else
                    {
                        GraphTokenException tokenException;
                        try
                        {
                            string responseContent = await response.Content.ReadAsStringAsync();
                            string responseStr = responseContent;
                            var serializer = new JavaScriptSerializer();
                            var result = serializer.Deserialize<Dictionary<string, object>>(responseStr);

                            string errorCode = "unknownError";
                            string suberror = null;
                            string description = null;
                            string claims = null;

                            if (result.ContainsKey("error"))
                            {
                                errorCode = result["error"].ToString();
                            }

                            if (result.ContainsKey("error_description"))
                            {
                                description = result["error_description"].ToString();
                            }

                            if (result.ContainsKey("claims"))
                            {
                                claims = result["claims"].ToString();
                            }

                            if (result.ContainsKey("suberror"))
                            {
                                suberror = result["suberror"].ToString();
                            }
                            tokenException = new GraphTokenException(errorCode, description, claims, suberror);
                        }
                        catch (Exception e)
                        {
                            throw (e);
                        }

                        throw tokenException;
                    }

                }

            }
        }
    }
}

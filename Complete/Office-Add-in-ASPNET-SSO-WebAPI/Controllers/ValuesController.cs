// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the root of the repo.

/* 
    This file provides controller methods to get data from MS Graph. 
*/

using Microsoft.Identity.Client;
using System.IdentityModel.Tokens;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web.Http;
using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
using Office_Add_in_ASPNET_SSO_WebAPI.Models;
using System.Web.Script.Serialization;
using System;
using System.Net;
using System.Net.Http;

namespace Office_Add_in_ASPNET_SSO_WebAPI.Controllers
{
    [Authorize]
    public class ValuesController : ApiController
    {
        // GET api/values
        public async Task<HttpResponseMessage> Get()
        {
            Dictionary<string, string> errorObj = new Dictionary<string, string>();

            // OWIN middleware validated the audience and issuer, but the scope must also be validated; must contain "access_as_user".
            string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
            if (addinScopes.Contains("access_as_user"))
            {
                // Get the raw token that the add-in page received from the Office host.
                var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext
                    as BootstrapContext;

                // Get the access token for MS Graph. 
                string[] graphScopes = { "Files.Read.All" };

                GraphToken result = null;
                try
                {
                    // The AcquireTokenOnBehalfOfAsync method will initiate the "on behalf of" flow
                    // with the Azure AD V2 endpoint.
                    result = await GraphApiHelper.AcquireTokenOnBehalfOfAsync(bootstrapContext.Token, graphScopes);
                }
                catch (GraphTokenException e)
                {
                    errorObj["claims"] = e.Claims;
                    errorObj["message"] = e.Message;
                    errorObj["errorCode"] = e.ErrorCode;
                    errorObj["suberror"] = e.SubError;
                    return SendErrorToClient(HttpStatusCode.Unauthorized, errorObj);
                }
                catch (Exception e)
                {
                    errorObj["errorCode"] = "unknown_error";
                    errorObj["message"] = e.Message;
                    return SendErrorToClient(HttpStatusCode.InternalServerError, errorObj);
                }

                // Get the names of files and folders in OneDrive for Business by using the Microsoft Graph API. Select only properties needed.
                var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");

                IEnumerable<OneDriveItem> filesResult;
                try
                {
                    filesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
                }

                // If the token is invalid, MS Graph sends a "401 Unauthorized" error with the code 
                // "InvalidAuthenticationToken". ASP.NET then throws a RuntimeBinderException. This
                // is also what happens when the token is expired, although MSAL should prevent that
                // from ever happening. In either case, the client should start the process over by 
                // re-calling getAccessTokenAsync. 
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException e)
                {
                    errorObj["errorCode"] = "invalid_graph_token";
                    errorObj["message"] = e.Message;
                    return SendErrorToClient(HttpStatusCode.Unauthorized, errorObj);
                }

                // The returned JSON includes OData metadata and eTags that the add-in does not use. 
                // Return to the client-side only the filenames.
                List<string> itemNames = new List<string>();
                foreach (OneDriveItem item in filesResult)
                {
                    itemNames.Add(item.Name);
                }

                var requestMessage = new HttpRequestMessage();
                requestMessage.SetConfiguration(new HttpConfiguration());
                var response = requestMessage.CreateResponse<List<string>>(HttpStatusCode.OK, itemNames);
                return response;

            }
            else
            {
                // The token from the client does not have "access_as_user" permission.
                errorObj["errorCode"] = "invalid_access_token";
                errorObj["message"] = "Missing access_as_user. Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.";
                return SendErrorToClient(HttpStatusCode.Unauthorized, errorObj);
            }

        }

        private HttpResponseMessage SendErrorToClient(HttpStatusCode statusCode, Dictionary<string, string> message)
        {
            var requestMessage = new HttpRequestMessage();
            requestMessage.SetConfiguration(new HttpConfiguration());
            HttpResponseMessage errorMessage = requestMessage.CreateResponse(statusCode, message);

            return errorMessage;
        }
        
        // GET api/values/5
        public string Get(int id)
        {
            return "value";
        }

        // POST api/values
        public void Post([FromBody]string value)
        {
        }

        // PUT api/values/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        public void Delete(int id)
        {
        }
    }
}

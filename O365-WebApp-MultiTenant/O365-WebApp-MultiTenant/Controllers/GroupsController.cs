using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;
using O365_WebApp_MultiTenant.Models;
using O365_WebApp_MultiTenant.Utils;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace O365_WebApp_MultiTenant.Controllers
{
    [Authorize]
    public class GroupsController : Controller
    {
        // GET: Contacts
        public async Task<ActionResult> Index()
        {
            var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            var userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            var tenantId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;

            AuthenticationContext authContext = new AuthenticationContext(string.Format("{0}/{1}",
                SettingsHelper.AuthorizationUri, tenantId), new ADALTokenCache(signInUserId));


            var authResult = await authContext.AcquireTokenSilentAsync("https://graph.microsoft.com/", 
                new ClientCredential(SettingsHelper.ClientId, SettingsHelper.AppKey),
                new UserIdentifier(userObjectId, UserIdentifierType.UniqueId));


            var url = "https://graph.microsoft.com/alpha/me/joinedGroups";

            var request = HttpWebRequest.CreateHttp(url);
            request.Method = "GET";
            request.Headers.Clear();
            request.Accept = "application/json, text/plain, */*";
            request.Headers["Authorization"] = "Bearer " + authResult.AccessToken;

            var response = request.GetResponse() as HttpWebResponse;


            var ResponseStr = new StreamReader(response.GetResponseStream()).ReadToEnd();


            return Content(ResponseStr);
        }
    }
}
# SP.MVC.HybridBoilerplate
Boilerplate for a SharePoint provider hosted application with API AND Web- Controller. The difficulty is to force the API Controller to execute code against the hostweb within the user context. The API controller has been extended to consume the session token from the web controller to access the app web!

# Test Setup
- SharePoint Online (June 2019)
- Visual Studio 2017

## References
- [Tutorial José Quinto](https://blog.josequinto.com/2016/09/05/how-to-provide-sharepointcontext-to-a-web-api-action-apicontroller-in-a-sharepoint-provider-hosted-app/)
- [Give your provider-hosted add-in the SharePoint look-and-feel](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/give-your-provider-hosted-add-in-the-sharepoint-look-and-feel)


## Change Steps API Controller Integration
1. Create a provider hosted application
2. Install Nuget Package for CSOM</br>
   [Nuget](https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM)</br>
   `Install-Package Microsoft.SharePointOnline.CSOM`</br>
   
   > Update all Nuget Packages caused some styling errors. So better do not updgrade the other Nuget Packages

3. Add API Controller (e.g Named Controllers\DataController.cs)
4. Add Filter (Filters\SharePointContextWebAPIFilterAttribute.cs)

   - ApiController class DOESN’T WORK with System.Web.Mvc.ActionFilterAttribute like the MVC Web does
   - ApiController WORKs with System.Web.Http.Filters.ActionFilterAttribute
   
   ```CSharp  
   using System;
   using System.Net;
   using System.Net.Http;
   using System.Web;
   using ActionFilterAttribute = System.Web.Http.Filters.ActionFilterAttribute;

   namespace SP.MVC.ReactBoilerplateWeb
   {
       public class SharePointContextWebAPIFilterAttribute : ActionFilterAttribute
       {
           public override void OnActionExecuting(System.Web.Http.Controllers.HttpActionContext actionContext)
           {
               if (actionContext == null)
               {
                   throw new ArgumentNullException("actionContext");
               }

               Uri redirectUrl;
               switch (SharePointContextProvider.CheckRedirectionStatus(HttpContext.Current, out redirectUrl))
               {
                   case RedirectionStatus.Ok:
                       return;
                   case RedirectionStatus.ShouldRedirect:
                       var response = actionContext.Request.CreateResponse(System.Net.HttpStatusCode.Redirect);
                       response.Headers.Add("Location", redirectUrl.AbsoluteUri);
                       actionContext.Response = response;
                       break;
                   case RedirectionStatus.CanNotRedirect:
                       actionContext.Response = actionContext.Request.CreateErrorResponse(HttpStatusCode.MethodNotAllowed, "Context couldn't be created: access denied");
                       break;
               }
           }
       }
   }
   ```


5. Decorate class with filter attribute in the Data Controller

   ```CSharp        
     [SharePointContextWebAPIFilter]
     // GET: api/Data/5
     public string Get(int id)
     { 
         return "value";
     } 
   ```
   
 6. Modify Global.asax
 
    ```CSharp
      using System.Web.Http;
      using System.Web.Mvc;
      using System.Web.Optimization;
      using System.Web.Routing;

      namespace SP.MVC.ReactBoilerplateWeb
      {
          public class MvcApplication : System.Web.HttpApplication
          {
              protected void Application_Start()
              {
                  GlobalConfiguration.Configure(WebApiConfig.Register);
                  AreaRegistration.RegisterAllAreas();
                  FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
                  RouteConfig.RegisterRoutes(RouteTable.Routes);
                  BundleConfig.RegisterBundles(BundleTable.Bundles);
              }
          }
      } 
    ```

 7. Change WebApiConfig.cs to enable Session State in Web API
 
    Web API (ApiController) is stateless component, which means that doesn’t have Session State.
   
     ```CSharp
     
      using System.Web;
      using System.Web.Http;
      using System.Web.Http.WebHost;
      using System.Web.Routing;
      using System.Web.SessionState;

      namespace MyWebApi
      {
          public static class WebApiConfig
          {
              public static void Register(HttpConfiguration config)
              {
                  RouteTable.Routes.MapHttpRoute(
                      name: "DefaultApi",
                      routeTemplate: "api/{controller}/{id}",
                      defaults: new { id = RouteParameter.Optional }
                  ).RouteHandler = new SessionRouteHandler();
              }

              public class SessionRouteHandler : IRouteHandler
              {
                  IHttpHandler IRouteHandler.GetHttpHandler(RequestContext requestContext)
                  {
                      return new SessionControllerHandler(requestContext.RouteData);
                  }
              }
              public class SessionControllerHandler : HttpControllerHandler, IRequiresSessionState
              {
                  public SessionControllerHandler(RouteData routeData)
                      : base(routeData)
                  { }
              }
          }
      }
     
     ```
     
 8. Call SharePoint from API Controller
 Finally call from the API Controller a SharePoint Context based execution --> Get Username
 
       ```CSharp
 
         // GET: api/Data/5
        [SharePointContextWebAPIFilter]
        [System.Web.Http.HttpGet]
        public string Get(int id)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext.Current);

            Microsoft.SharePoint.Client.User spUser = null;
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    spUser = clientContext.Web.CurrentUser;

                    clientContext.Load(spUser, user => user.Title);

                    clientContext.ExecuteQuery();
                }
            }
            return "user: " + spUser.Title + " - " + id;
        }

 
      ```

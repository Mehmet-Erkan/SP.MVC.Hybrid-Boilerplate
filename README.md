# SP.MVC.ReactBoilerplate
Boilerplate for a provider hosted SharePoint application with React, Typescript, API Controller and PnP. The difficulty is to force the API Controller to execute code against the hostweb within the user context. 

# Test Setup
- SharePoint Online (June 2019)
- Visual Studio 2017


## Installation
1. Create a provider hosted application
2. Install Nuget Package for CSOM</br>
   [Nuget](https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM)</br>
   `Install-Package Microsoft.SharePointOnline.CSOM`</br>
   
   > Update all Nuget Packages caused some styling errors. So better do not updgrade the other Nuget Packages

3. Add API Controller (e.g Named Controllers\DataController.cs)
4. Add Filter (Filters\SharePointContextWebAPIFilterAttribute.cs)

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

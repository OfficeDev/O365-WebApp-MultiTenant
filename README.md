Office 365 APIs MultiTenant Web Application
============================================

This sample shows how to build an MVC web application that uses Azure AD for sign-in using the OpenID Connect protocol, and then calls a Office 365 API under the signed-in user's identity using tokens obtained via OAuth 2.0. This sample uses the OpenID Connect ASP.Net OWIN middleware and ADAL .Net.

## How to Run this Sample
To run this sample, you need:

1. Visual Studio 2013
2. [Office 365 API Tools for Visual Studio 2013](https://visualstudiogallery.msdn.microsoft.com/7e947621-ef93-4de7-93d3-d796c43ba34f)
3. [Office 365 Developer Subscription](https://portal.office.com/Signup/Signup.aspx?OfferId=6881A1CB-F4EB-4db3-9F18-388898DAF510&DL=DEVELOPERPACK&ali=1)

## Step 1: Clone or download this repository
From your Git Shell or command line:

`git clone https://github.com/OfficeDev/O365-WebApp-MultiTenant.git`

## Step 2: Build the Project
1. Open the project in Visual Studio 2013.
2. Simply Build the project to restore NuGet packages.
3. Ignore any build errors for now as we will configure the project in the next steps.

## Step 3: Configure the sample
Once downloaded, open the sample in Visual Studio.

### Register Azure AD application to consume Office 365 APIs
Office 365 applications use Azure Active Directory (Azure AD) to authenticate and authorize users and applications respectively. All users, application registrations, permissions are stored in Azure AD.

Using the Office 365 API Tool for Visual Studio you can configure your web application to consume Office 365 APIs. 

1. In the Solution Explorer window, **right click your project -> Add -> Connected Service**.
2. A Services Manager dialog box will appear. Choose **Office 365 -> Office 365 API** and click **Register your app**.
3. On the sign-in dialog box, enter the username and password for your Office 365 tenant. 
4. After you're signed in, you will see a list of all the services. 
5. Initially, no permissions will be selected, as the app is not registered to consume any services yet.
6. Select **Users and Groups** and then click **Permissions**
7. In the **Users and Groups Permissions** dialog, select **Enable sign-on and read users** profiles' and click **Apply**
8. Select **Contacts** and then click **Permissions**
9. In the **Contacts Permissions** dialog, select **Read users' contacts** and click **Apply**
10. Click on **App Properties** and select **Multiple Organizations** to make this app multi-tenant.
10. Click **Ok**

After clicking OK in the Services Manager dialog box, Office 365 client libraries (in the form of NuGet packages) for connecting to Office 365 APIs will be added to your project. 

In this process, Office 365 API tool registered an Azure AD Application in the Office 365 tenant that you signed in the wizard and added the Azure AD application details to web.config. 

### Step 4: Build and Debug your web application
Now you are ready for a test run. Hit F5 to test the app.

### Quick Look at the SSO Authentication Code
The authentication startup class, **App_Start/Startup.Auth.cs** in the project contains the startup logic for Azure AD authentication. 

The project implements a simple ADAL token cache **NaiveSessionCache** that uses the ASP.Net session to store and retrieve tokens for the current user. **As it name suggests, it is very naive and is not recommended for production use**.A more persistent cache such as database is recommended for production use.

### Sign In and Sign Out Controls
The sign in and sign out controls are already added to the views. You can find them under **Views\Shared** folder.
1. **_LoginPartial.cshtml** is the partial view that renders the Sign In and Sign Out actions.
2. **_LoginPartial.cshtml** is then rendered in _Layout.cshtml
3. The **AccountController.cs** has the required methods for sign in and sign out.

### Requiring authentication to access controllers
Applying **Authorize** attribute to all controllers in your project will require the user to be authenticated before accessing these controllers. To allow the controller to be accessed anonymously, remove this attribute from the controller. If you want to set the permissions at a more granular level, apply the attribute to each method that requires authorization instead of applying it to the controller class.

### Write Code to call Office 365 APIs
You can now write code to call an Office 365 API in your web application. You can apply the Autorize attribute to the desired controller or the method in which you wish to call Office 365 API.

**ContactsController.cs** describes how to interact with the Office 365 API discovery service, get the endpoint URI and resource Id for Outlook Services and then query users' contacts.





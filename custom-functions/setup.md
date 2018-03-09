_[Workshop home](../index.md)_  >  _[Custom functions](index.md)_ > _Setup_


# Setup

These instructions will let you set up your computer to preview custom functions in Excel. You'll need a Windows PC for now - the preview hasn't been released for Mac or other platforms yet.

## Office Insider Program
Join the [Office Insider program](https://products.office.com/en-us/office-insider?tab=Windows-Desktop) to install the version of Excel 2016 that's required for custom functions on your computer (must also be version 16.8711 or later). You must choose the "Insider" channel for the custom functions preview to work.

You can check to see whether you're already on the Insiders channel by clicking File > Account, then look for the channel next to "About Excel".

If you're using an organizational subscription (enterprise or EDU), then you'll need to follow [these longer instructions](https://support.office.com/en-us/article/how-office-365-commercial-customers-can-get-early-access-to-new-office-2016-features-4dd8ba40-73c0-4468-b778-c7b744d03ead?ui=en-US&rs=en-US&ad=US) to download the Office 2016 Deployment tool, manually modify a configuration file, and run the tool (sorry!). Otherwise it's easy to opt-in to Insiders: just go to File > Account > Office Insider and select "Get early access to new releases of Office".

## Install the sample template

We've prepared a sample add-in with several custom functions. Download the [Excel-Custom-Functions](https://github.com/saunders77/Excel-Custom-Functions) code or clone the repo. There are 3 important files: the JS, HTML, and XML.

Host the files on a web server. You can follow the steps below to host and edit them quickly via Azure Portal, or you can choose any alternate method like hosting on some other server or install IIS and host locally on your machine.
  1. Go to [https://portal.azure.com](https://portal.azure.com) and sign in with a Microsoft account.
  2. Create a resource > Web App
  3. Sign up for a subscription if you don't already have one (requires a credit card for verification but you won't be charged unless you explicitly convert to a paid service, which isn't required for this walkthrough).
  4. Choose a name for your web app. Other settings can be left on their defaults. Choose "pin to dashboard" and click Create
  5. Wait for your web app to deploy (around 30-90 seconds)
  6. In your Web App resource, scroll down and choose "App Service Editor (Preview)" under Development Tools
  7. In WWWROOT, choose "Upload Files" and upload the HTML and JS files in the sample. (Alternatively, if you prefer your own editor/client, you can deploy to the Azure Web App with git.)
  8. Create a new folder on your main drive and share it with yourself (Properties > Sharing > Share)
  9. Copy the XML manifest file from the sample to your shared folder.
  10. If you've installed XML manifest files before on this computer, modify the name of the file and the ID element (you can use Notepad and change a couple random digits to make sure it doesn't conflict)
  11. In the XML manifest, modify the INSERT-URL-HERE to the URL you chose for your web app (eg. "mycustomfunctions.azurewebsites.net")
  12. Launch Excel and configure side-loading at File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs. Enter the path of the shared folder you created (in shared-path format starting with \\). Click "Add catalog" and check the "Show in Menu" box. Click OK and restart Excel
  13. Insert > My Add-ins > Shared Folder > Custom Functions Sample > Insert
  14. Type =CONTOSO to get a list of available functions from the sample.

You can debug with Visual Studio: first enable script debugging in Internet Explorer (go to Internet Options > Advanced > uncheck "Disable Script Debugging" for IE and others). Then during an Excel session with custom functions running, choose "Attach to Process" in Visual Studio and select the iexplore process that corresponds to the add-in (should say "Script" in the type field). Alternatively, add "log" statements to your JavaScript to write them to the Excel DEBUG function. 




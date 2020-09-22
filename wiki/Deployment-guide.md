# Prerequisites

To begin, you will need:
* A Microsoft 365 subscription
* A team with the users who will be sending Checklists using this app. (You can add and remove team members later!)  
* A copy of the Checklist app GitHub repo (https://github.com/OfficeDev/microsoft-teams-checklist-app) 
* Install Node.js (using https://nodejs.org/en/download/) locally on your machine. 


# Step 1: Create your Checklist app

To create the Teams Checklist app package:
1. Make sure you have cloned the app repository locally.
1. Open the actionManifest.json file in a text editor.
1. Change the placeholder fields in the manifest to values appropriate for your organization. 
    * packageID - A unique identifier for this app in reverse domain notation. E.g: com.contoso.checklistapp. (Max length: 64) 
    * developer.[]()name ([What's this?](https://docs.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema#developer))
    * developer.websiteUrl
    * developer.privacyUrl
    * developer.termsOfUseUrl

Note: Make sure you do not change to file structure of the app package, with no new nested folders.  


# Step 2: Deploy app to your organisation

1. Install Node.js (using https://nodejs.org/en/download/) locally on your machine.
1. Open Command Line on your machine.
1. Navigate to the app package folder with the name `microsoft-teams-checklist-app`.
1. Run the following command to download all the dependent files mentioned in package.json file of the app package.

    **```npm install```**
1. Once the dependent files are downloaded, run the below command to deploy the app package to your Microsoft 365 subscription. When prompted, log in to your AAD account.  

    **```npm run deploy```**
1. An AAD custom app, Bot are programmatically created in your tenant to power the Checklist message extension app in Teams.
1. Once run, a Checklist Teams app zip file is generated under output folder in the same directory as your cloned app repository locally with the name `microsoft-teams-checklist-upload.zip`.


# Step 3: Run the app in Microsoft Teams

If your tenant has sideloading apps enabled, you can install your app by following the instructions [here](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/apps/apps-upload#load-your-package-into-teams).

You can also upload it to your tenant's app catalog, so that it can be available for everyone in your tenant to install. See [here](https://docs.microsoft.com/en-us/microsoftteams/tenant-apps-catalog-teams).

Upload the generated Checklist Teams app zip file (the `microsoft-teams-checklist-upload.zip`) under output folder in the same directory as your cloned app repository locally to your channel, chat, or tenant’s app catalog. 


# Step 4: Update your Checklist Teams app

If you want to update the existing Checklist Teams app with latest functionality -
1. Make sure you have cloned the latest app repository locally.
1. Open the `actionManifest.json` file in a text editor.
    * Change the placeholder fields (`packageID`, `developer.name`, `developer.websiteUrl`, `developer.privacyUrl`, `developer.termsOfUseUrl`) in the manifest with existing values in your Checklist Teams app. 
    * Update the `version` field in the manifest. Make sure latest version number is higher than previous version number.  
1. Run the below command to update your Checklist Teams app with the latest bits of code. When prompted, log in using your AAD account. 
    
    **```npm run update-teams-app```**
1. Your Checklist app on Teams automatically gets updated to the latest version. 

# Troubleshooting

Please see our [Troubleshooting](Troubleshooting.md) page.

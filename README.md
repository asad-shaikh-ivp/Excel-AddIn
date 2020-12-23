# IVP Research Management Frameowrk Excel Add-In

## Summary

This excel add-in is based on Microsoft Office Add-in JavaScript API , as a single-page application that does the following:
1. Connects to Azure AD for user authentication (using OAuth 2.0 authorization framework) and gets the Authentication token
2. Scrapes the open excel sheet to scrape the Company Model Data as configured for the C4 company models
3. Uses the Authentication token gotten in step 1 to connect to the API service to upload the company model data


## Prerequisites

To run this code, the following are required.

* [Node and npm](https://nodejs.org/en/), version 10.15.3 or later (npm version 6.2.0 or later).

* TypeScript version 3.1.6 or later.  - added in devDependencies

* An AzureAD account that is configured in both this Extension and the Backend API where the scraped excel data is to be uploaded

## Solution


## Build and run the solution

### Configure the solution
 In the code editor, open the `/login/login.ts` file in the project. Near the top is a configuration property called `clientId`. Replace the placeholder value with the application ID you copied from the registration. Save and close the file.

To use self signed certificates for development, follow the instructions at [SSL](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md#method-2-manually-install-the-certificate) to trust a certificate. Use method 2.

Run the command `npm install`.

### Run the solution

Run the command `npm start`.


## Scenario: Custom function batching

The code for custom functions runtime is placed under src/functions/functions.ts

In this scenario the CFRMS custom functions call a remote service. To reduce network round trips we batch all the calls and send them in a single call to the web service. This is ideal when the spreadsheet is recalculated. For example, if someone used custom function in 100 cells in a spreadsheet, and then recalculates the spreadsheet, custom function would run 100 times and make 100 network calls. By using this batching pattern, the calls can be combined to make all 100 calculations in a single network call.

### Key parts of Custom Functions
Instead of performing the calculation, each of them calls a `_pushOperation` function to push the operation into a batch queue to be passed to a web service.

#### Batching the operation
The `_pushOperation` function pushes each operation into a _batch variable. It schedules the batch call to be made.

#### Making the remote request
The `_makeRemoteRequest` function prepares the batch request and passes it to the `_fetchFromRemoteService` function. If you are adapting this code to your own solution you need to modify `_makeRemoteRequest` to actually call your remote service.

#### The remote service
The `_fetchFromRemoteService` function processes the batch of operations, fetches data from server, and then returns the results.
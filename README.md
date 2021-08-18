Syncs Xero Bank Statements with Google Sheets

This uses the Xero private application connection

1. Create a blank Google Sheet with two sheets: Statement and Technical
2. In Tools-Script editor create a new file and paste code.gs there
3. In Xero Developer create a new app
Integration type-Web App
Company or application URL-https://test.com
OAuth 2.0 redirect URIs-http://localhost:5000/callback
Create App
4. In code.gs change these
var XERO_CLIENT_ID = 'EMPTY'; // "Client id" from your app at https://developer.xero.com/app/manage
var XERO_CLIENT_SECRET = 'EMPTY'; // "Client secret" from your app at https://developer.xero.com/app/manage
5. Run updateStatement script
On first execution it gets the authorization URL
Paste the authorization URL in your browser
Authorize the app to connect
Copy the value from the URL http://localhost:5000/callback?code=11111111111111111111&scope=...
var XERO_AUTHORIZATION_CODE = '11111111111111111111'; // get this from the URL you were redirected to
6. Run updateStatement script
On second execution it pulls the bankaccountID and also the last transactions of the bank statement.
If you want to define the bank account add the bankAccountID to "Technical Sheet" Cell B1.
6. Schedule a trigger for updateStatement
Created by emphasoft.com
Syncs Xero Bank Statements with Google Sheets

This uses the Xero private application connection

1. Create a blank Google Sheet with two sheets: Statement and Technical
2. In Tools-Script editor create 3 files and put the .gs files there.
3. In code.gs change your Consumer Key, User Agent (Xero Application Name) and PEM Key.
var XERO_CONSUMER_KEY = "AAAAAAAA";
var XERO_USER_AGENT = "TestApp";
var XERO_PEM_KEY = "-----BEGIN RSA PRIVATE KEY----- AAAAAAAA -----END RSA PRIVATE KEY-----";
4. Run updateStatement script
On first execution it pulls the bankaccountID and also the last transactions of the bank statement.
If you want to define the bank account add the bankAccountID to "Technical Sheet" Cell B1.
5. Schedule a trigger for updateStatement

I would like to thank the Lead Contributor of this project, Evgeny Rozhnov. You can reach him at info@deliveryweb.ru for similar projects if you need
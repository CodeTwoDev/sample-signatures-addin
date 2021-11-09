# Sample signatures add-in
This is a simple Outlook Web add-in that uses event-based activation to insert a static email signature and a disclaimer depending on message recipients.
It works like this: when you create a new message in Outlook, the signature and disclaimer are downloaded and cached in the SessionData object.
When you add an internal recipient, the add-in inserts the signature.
If you add an external recipient, the add-in inserts the signature and the disclaimer.

## Learn More ##
To learn more about Office Add-ins, see the [Office Add-ins documentation](https://aka.ms/office-add-ins-docs) and check [other samples](https://github.com/OfficeDev/PnP-OfficeAddins).
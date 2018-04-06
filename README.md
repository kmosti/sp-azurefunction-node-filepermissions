### Handle security and versioning for drop off library

This is a node.js application intended for Azure Function which will go through the drop off library, then set sharing permissions for each file to the sender of the document.
The document will be sent through MS Flow, where one of the item properties will be Sendername containing the email of the sender.
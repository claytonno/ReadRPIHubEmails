# ReadRPIHubEmails


This is a simple tool I created to take images that come in via email through Outlook, pass them through YOLO/ONNX to determine
if there are any objects found within to be labeled "person" and if so to save them to a "Saved" folder. 

The purpose of this for me was that I use a motion capture sercurity system called MotionEyeOs on Raspberry PI, that I found impossible to perfectly fine 
tune to motion due to trees and other artifacts causing false positives. This tool will automatically move the emails to the deleted folder if 
there are no matches, and it will move the images to a "saved" folder locally if there are any matches, with a time stamp. 

This has the additional benefit of marking all the emails as "Read" by default, an option not available in IOS' Mail App, so when I
check my phone I never have any mail notifications from the mailbox where these emails are sent. I also do not have to forward any
ports on my router for this system to work since the notification is driven by email.

The model itself is not perfect and can have some strange matches but I'm working on improving it.

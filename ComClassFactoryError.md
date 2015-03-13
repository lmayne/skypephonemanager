# Error message "Retrieving the COM class factory failed" on startup #

## SYMPTOMS ##
Some users (the pattern appears to be people using non-english versions of Windows) are receiving the following error when starting the SkypePhone Manager:

**_Retrieving the COM class factory for component with CLSID {830690FC-BF2F-47A6-AC2D-330BCB402664} failed_**

## CAUSE ##
This problem is caused by the deployed Skype interop DLL not being compatible with your version of Windows

## RESOLUTION ##
To fix this issue:

  1. Make sure you have the latest version of Skype. Download and install it from www.skype.com
  1. Download the Skype4COM DLL from: http://skypephonemanager.googlecode.com/svn/trunk/SkypePhoneManager/Skype4COM.dll and put it in the same folder as the other SkypePhone Manager files (e.g. C:\SkypePhoneManager, or wherever you unpacked it)
  1. Open a command prompt (Start / Run / cmd)
  1. In the command prompt cd to the above directory (cd C:\SkypePhoneManager)
  1. Run the following command: c:\Windows\System32\regsvr32.exe Skype4COM.dll

You should then be able to start the application without the exception occuring.
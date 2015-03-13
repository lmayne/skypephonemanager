# Important! #
**You will need version 2.0 of the .NET framework installed to run this application.**

# Installation #

In order to set up the application, you will need a PC running constantly (or for the duration of your usage) with Skype installed on it.

  1. Create two Skype accounts, one for your PC and one for your phone (we will refer to these accounts as pcuser and mobileuser respectively)
  1. Log into Skype on the PC as pcuser, and add the mobile account to the contact list
  1. Buy a SkypeIn number and some SkypeOut credit for pcuser
  1. Set up forwarding for pcuser to a landline number you think you might call a lot (like your home phone), using Tools / Call Forwarding
  1. Log into your Skypephone as mobileuser and make sure you can make Skype calls back and forth between the PC and the phone
  1. Unzip the file you downloaded from this site to somewhere on your PC, like C:\SkypePhoneManager
  1. In the folder you unzipped to, edit the SkypePhoneManager.exe.config file and enter your mobile phone username in the appropriate space (e.g. mobileuser)
  1. Run SkypePhoneManager.exe, and select the option to always allow the application to use Skype. Leave the application running.

# Common Problems #
There seem to be two main problems that people run into when installing the application:
  1. Not extracting all the files. A lot of people assume the exe is an installer and try double-clicking it in Winzip or similar. You need to extract all the files in the zip to a folder before configuring and running the exe.
  1. Unable to open the config file. This is a plain text document which should be opened using notepad or wordpad, NOT Microsoft Word.
  1. Changing both the key and the value in the config file. A lot of people change both the setting name and the value instead of just the value. For example, if your username is fredbloggs then your config line should look like this:

```
<add key="SkypeMobileUsername" value="fredbloggs"/>
```
NOT:
```
<add key="fredbloggs" value="fredbloggs"/>
```
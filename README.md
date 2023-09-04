# Email AI Reader (EAR)

Scripts to monitor and manage a common email box, making suggestions to support the team.

## Why would I want to do this?

Many organisations use a simple Outlook mailbox to cordinate client responses.

Outlook allows you to do this ... but ...

Since .... it allows you to do the following things:

1. Report likely things that go wrong
1. Allows you to backup Tasks if you accidentally delete them in Outlook.
1. By saving to OneDrive, Google Drive or similar, it makes one tasklist available across multiple devices.

## What this Python Script does

Given a typical Outlook share folder like this:

![Outlook Tasks Screenshot](images/outlook-tasks.png)

The script generates a report ...

![Excel Tasks Screenshot](images/excel-tasks.png)

## Safety first

Safety first, we treat Outlook as the 'gold' copy.

* We never delete any data from outlook, and can always (re)create reports from this data.
* We never send any email - any suggestions are saved as drafts, or tagged emails.
* Everything is 'Human in the loop' - a real person needs to decide to send email based on a suggestion.

## Getting Started

1. Make sure you have Outlook on your machine.
1. [Install Python](https://www.python.org/downloads/) on your machine.
1. Make sure you have the required libraries - typically this will be something like ``pip install pandas openpyxl pywin32`` in a terminal.
1. Run the script in a terminal using a command similar to ``python capture.py`` (Note capture.bat will create directories if needed)
   * By Default - the script will look for the templates and outputs in the same directory as it is run. Log files and backups will also by placed in this directory.
   * capture.bat is there for convenience.

## Modifying the Script

The main configuration is in settings.py with comments to allow you to easily edit.

The comments in the ``capture.py`` script should make it pretty clear what is going on. The Excel file names, the log file names and backups are all set as constants at the top of the file (e.g. if you want change the Excel task file location).

If you want to extract / upload different properties from the Outlook tasks (e.g. percent complete), the pattern should be very familiar. The names used in code may differ in the Outlook object model from what you use in the Outlook Desktop interface. A link is given below to the Microsoft reference to help you.

## More Technical Information

The approach taken is to use the API provided by Outlook's COM model, rather than the newer Microsoft Graph API. The reason for this is that (currently) not all Task information is exposed by the Graph API - it appears to be limited to the few fields used in the Microsoft Todo app.

There is a lot of information on the Web describing PyWin32 - the library used to connect Python to Windows Applications like Outlook. There is less information on the Object Model within Outlook - and mostly it is intended for VBA and C# users (athough the method calls and params are very similar). Some good starting points used in creating this script:

* [Microsoft Docs describing the Outlook Com Object model](https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mapifolder?view=outlook-pia)

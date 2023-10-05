Note
----
This repository is for history. For now, the issue described here is not reproducible.

Prerequisites
-------------
Addin code is hosted on: https://danilkorotenko.github.io/OutlookAddin/index.html

[Manifest](/manifest.xml)

Issue Description
-----------------
Outlook on-send addin is not always being invoked in the Outlook for Mac app.

Steps To Reproduce
------------------
1. Install manifest: [Manifest](/manifest.xml)
2. Create a new mail message, with "To:", "Subject:" and some content.
3. Press the "Send" button.
4. The "Send" button becomes disabled for a moment.
5. A separate message window appears with the "Message blocked." notification.
6. Click the "Send" button again several times.

Actual Result
-------------
The message is sent.

Expected result
---------------
The message should always be blocked to send.

Additional info
---------------
This issue is absent in Outlook On the Web (https://outlook.com)

This issue is absent in Outlook for Windows version: 2108 (build 14332.20447).

Hardware
--------
MacBook Air M1, 2020

Software
--------
macOS 13.2.1

Outlook for Mac 16.73

Screenshots
-----------
![](/Screenshot3.png)
![](/Screenshot1.png)
![](/Screenshot2.png)
![](/Screenshot4.png)

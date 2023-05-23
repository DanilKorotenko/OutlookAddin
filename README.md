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
3. Press "Send" button.
4. The "Send" button become disabled for a moment.
5. A separate message window appear with "Message blocked." notification.
6. Click "Send" a few times.

Actual Result
-------------
The message is sent.

Expected result
---------------
The message should always be blocked to send.

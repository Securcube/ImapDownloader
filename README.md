# SecurCube IMAP Downloader

[![N|Solid](http://securcube.net/wp-content/uploads/2016/01/logo2.png)](http://securcube.net/)

SecurCube ImapDownloader is a free forensic tool that allows to clone an imap account in folders and emails (.eml).
At the end of the process you will get a zip file with all mails organized in forders correlated with hashes.

Features
----

- Fast connection parameters selection for:
  - Gmail
  - iCloud
  - Outlook.com
  - Yahoo!
  - AOL
  - AT&amp;T
  - MSN
- Multi-thread folders downloads to speedup the process
- Support download resume in case of lost connection, errors, or whatever
- Selective folders download
- Also important email details are saved, such as:
    - Email ID (IMAP folder progressive ID)
    - MessageId
    - Email date
    - Email Not Readed status (marked with '_N_' in the filename)
- PCAP acquisition of all imap commands
	
Emails are named like:
<ID>(\_N)\_<MessageID>.eml
examples:
18_2723D5A69F7749030FEC@AMSPRD40.064d.mgd.msft.net.eml      <- Seen message
25_N_141015162414KJ.05013@mscreator39.fagms.de.eml          <- Not Seen Message (or marked as Unseen)

This is a sample of the log file:

```sh
LOGIN PARAMS
Host name: imap.gmail.com (resolved ip = 111.222.333.444) 
Port: 993
UseSSL: True
User name: dummy@gmail.com
User password: ------

Folder: [Gmail]		0 emails
Folder: [Gmail]/Bin		83 emails
Folder: [Gmail]/Drafts		0 emails
Folder: [Gmail]/Spam		122 emails
Folder: [Gmail]/Starred		0 emails
Folder: [Gmail]/Important		3456 emails
Folder: [Gmail]/Sent Mail		80 emails
Folder: INBOX		15262 emails

Total emails: 22472

Startd at 17/07/2017 09:45:45 UTC
End at 17/07/2017 11:00:36 UTC

Export file : D:\ExportImap\dummy@gmail.com.zip
MD5 : 6AF9632A420D610EA8496225114F6B8E
SHA1 : 196310408C3AF3A3673DB36B0475D0D242337AD9
```

Prerequisites
----
- [Microsoft .NET Framework 4.6](https://www.microsoft.com/download/details.aspx?id=48130)
- [WinPcap](https://www.winpcap.org/) - Optional

Downloads
----
## [Click here to download the latest release](https://github.com/Securcube/ImapDownloader/releases)

Known issue
----
When Windows crashes sometime the zip file is corrupt and no resume is supported. By the way you can open and fix the file with 7-zip.

Bug reports or new requests
----
Please go to http://devfarm.it/forums/forum/imapdownloader/
We will answer as soon as we get a chance.
 

Development
----
Want to contribute? Great!

License
----
[GPL-3.0](https://choosealicense.com/licenses/gpl-3.0/)

**Free Software, Hell Yeah!**

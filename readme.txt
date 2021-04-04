Thank you for using this software.
----------------------------------

This software is developed by
                                          SRUBS,
                                     C-126, 5th Cross,
                            Thillai Nagar, Trichy - 620018. India
                                     Phone. 91 431 765876
                    Email: srubs@vsnl.com. Visit: www.trichy-online.com

                               Send your comments - T. Rajesh


Instructions
------------


Double click or Press Esc to clear msg body text box. To select the records to send the mails,
put * in the select column of the grid. To unselect the records, delete * from the select column
of grid and press space bar. Only the selected  rows are processed for the bulk mail. Type body
message in the text box. Before sending bulk mails select attachment by Attachment button and it
accepts only one attachment.

To send bulk mails:  * Press Attachment button to select the attachment
                     * Press the Session On button to activate the mailing session
                     * Press Send Mails to send bulk mails

          F1 - Help F2 - About F3 - Attachment F4 - Session on F5 - Send F6 - Exit

This software uses the MAPI to send the mails. Run this software correctly use cdo.dll file to
be registered first. The cdo.dll can be downloaded from the MicroSoft site at

	CDO for Windows 9x

		http://www.microsoft.com/exchange/downloads/CDO_win95.exe

	CDO for Windows NT

		http://www.microsoft.com/exchange/downloads/CDO_WinNT.exe

Installing CDO
-------------------

1.	Copy the appropriate CDO.DLL file to the Windows\System32 Directory.  

	Note: 
		There is a two versions of the CDO.DLL: one for 
		Windows 95/Windows 98 and another for Windows NT. 
 		Make sure you copy the correct DLL to your machine.

2.	Register the CDO.DLL by executing REGSVR32.EXE CDO.DLL.

If Active Messaging is installed, it is not necessary to unregister OLEMSG32.DLL 
or ACTMSG.DLL, although it is good practice to do so. Also, once CDO.DLL is 
installed properly, neither OLEMSG32.DLL nor ACTMSG32.DLL are required for CDO 1.21 
to function.

------------------------
Overview of CDO
------------------------

Microsoft Collaboration Data Objects (CDO) is a technology for building messaging 
and collaboration applications. The current version of CDO is 1.21. 
It is designed to simplify the creation of applications with messaging 
functionality, and to add messaging functionality to existing applications. 
For example, CDO and Active Server Pages enable you to add scripts to a Web site 
to provide support for creating, sending, and receiving e-mail as well as 
participating in discussions and other public folder applications.

CDO does not represent a new messaging model, but rather an additional scripting 
interface to the Messaging Application Programming Interface (MAPI) model. 
CDO exposes programmable messaging objects (including folders, messages, 
recipient addresses, attachments, and other messaging components), which are 
extensions to the programmable objects offered as part of Microsoft Visual Basic(r),
such as forms and controls.

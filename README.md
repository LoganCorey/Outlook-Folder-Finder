# Outlook-Folder-Finder

Purpose of program:

This programs intended use case is finding folders in outlook if "someone" is particularly fond of nesting folders making them difficult to use.  For instance, if you're wondering where a folder "microsoft" is located, you could type in the command (once compield) flookup microsoft and it would return mail_box/foo/bar/microsoft. Note: an outlook session must currently be open!

Compiling:
A reference to Outlook 2016 object must be made in order to access Outlook.Application

Commands:

Find folder from default mailbox
flookup {folder_name}

example output:
  default_mail_box/foo/bar/microsoft
  
Find folder in other mailbox
flookup -i {mail_box} {folder}

example output:
  mail_box/foo/bar/microsoft
  
Set the default mailbox:
flookup -d {inbox_name}

Get the default mailbox:
flookup -d
example output: 
  foo@outlook.com

Get all available mailboxes:
flookup -s

example output:
  foo@outlook.com
  bar@outlook.com

Help:
flookup



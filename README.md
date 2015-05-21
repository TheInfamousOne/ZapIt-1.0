# ZapIt-1.0
My first program I created to Install, Uninstall and Zap the software out of the registry.  

Zapping the registry will not remove installation directory or files.  The only leverage this program gives is to remove
software that can not be removed so you can try to reinstall it.  This uses Microsofts MSIZap.exe.

Please remember to download and use the WindowsInstaller.exe and add it in the same root directory with ZapIt. 

What is ZapIt doing?
ZapIt will connect to a TERMID or NetBios name and copy the WindowsInstaller.exe program in C:\WSMGMT\Bin and then execute. Next
it will return all the software installed per-machine or per-user and grab the name and product guid. Just highlight the 
software you want to uninstall or Zap out of the registry.  You can figure out the rest.

Ty Stallard 
University of Michigan
MCIT-DES-EDEM
tyrones@med.umich.edu
tystallard@gmail.com
Published on GITHUT ON 05/21/2015
Developed on 02/18/2015


# Electronic-invoice-document-processing
######### The Project #########

The idea of this project is to process electronic invoices in PDF file format and to prepare them for further document processing in e.g. SAP or other document management solution. What exactly this project is doing is that
it saves all the PDF attachments of an e-mail to a folder and then merge them into one file and classify the file based on the sender's mail address to any of the predefined categories and for the sake of the document manageability 
a standard bar-code is being placed at the top of the output file and some compression is being applied to save space, but quality is also taken in consideration, so at the end of the processing all elements of the final output file are readable.
The project can also process manually created (e.g. scanned or printed through a PDF printer files), but please take in consideration the if the file has been scanned, the input file's resolution needs to be quite high, preferably 600 DPI
in order to ensure that the output of the processing will be readable and not artefacts/scrambled text is present on the final file.

######### LICENSE #########

This project is licensed under the MIT license, therefore is free to use, distribute and modify according to the user's needs. For more information chek the MIT_License.txt file.


######### INSTALLATION #########
1) OS - basically any Linux distro is suitable, but my tests showed that openSUSE provides better font rendering than Debian/Ubuntu (due to tight deadlines I did not have the time to investigate why). Due to clients' infrastructure 
currently 32 bit version is being deployed, but in general the OS architecture is not of importance. Due to its LTS openSUSE 13.1 is being used.

2) Required software:

2.1) GhostScript, GhostScript fonts - currently version 9.19 is being deployed, no plans for upgrade at this point

2.2) ImageMagick - currently used version 6.9.3 - Q16

2.3) GraphicsMagick - currently used version 1.3.20

2.4) libtiff5 -  currently used version 4.06

2.5) JPEG -  currently used version libjpeg62 - 62.2.0; libjpeg8 - 8.1.2; libjpeg-turbo - 1.5.0

2.6) mingetty -  currently used version 1.0.8s

######### ENVIRENMONT #########
1) Set the VM to see the shared folders and make sure it has full read and write access
2) Set the VM's NAT to point to the SSH daemon for troubleshooting
3) Set mgetty to autologin (on openSUSE edit /etc/systemd/system/getty.target.wants/getty@tty1.service and edit the line ExecStart=-/sbin/agetty --autologin scanpot --noclear %I $TERM)
5) Copy the folders.sh, instanceV6_MIFF.sh and variables.incl in the HOME folder and run the folders.sh script to recreate the internal folders used for the processing and to copy the scripts to the newly created workfolder
6) Set the instanceV6_MIFF.sh script to start at user's login
7) Copy the contents of Outlook_Script_With_Sleep.txt to the operator's Outlook VBA console and restart Outlook

######### Bar-code Generation #########
I used Zint (https://sourceforge.net/projects/zint/) to generate the needed bar-codes, there are lots of FOSS solutions to do this.

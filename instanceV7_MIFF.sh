#!/bin/bash

#sleep 60s
username="$(id -u -n)"

#$username@$HOSTNAME:~/workfolder$ cat instanceV7_MIFF.sh



source variables_V7.incl
today=`date '+%d%m%Y_%H%M%S'`;

function check_folder {

############################ cd Input Invoicre folder
clear="/home/$username/"
cd $3
#echo $3
rm Thumbs.db
#mv *.pdf $6

#Check if invoice.txt exists, should ensure that all files have been saved to the server
if [[ -f "invoice.txt" ]]; then
echo "Begin program execution..."
sleep 7s
mv *.* $6

cd $6
#echo $6
ls *.pdf | grep -oe '[0-9]\{4,25\}' | grep -v '[0-9]\{14\}' | sort | uniq > invoices.txt # get the number of invoices
cat invoices.txt
#inv_num=`wc -l "invoices.txt" | grep -oe '[0-9]'` #count the number of invoices
inv_num=`cat invoices.txt`
#check for invoice and application and merge the files
echo "inv_num is:" $inv_num
filename="invoices.txt"

echo "filename is: " `cat $filename`

#IFS=$'\n'
#process = false
for ((i =0; i <= inv_num; i++))
 do

  for invNum in $(cat invoices.txt)
  do
         cd $6
#        echo "$6 is: " $6
         echo "Entering FOR loop"
	     echo "invNum is: " $invNum
         ls -m1  | grep "$invNum" > tmpFile.txt
         echo "tmpFile.txt is: " `cat tmpFile.txt`
         
         countInvAndApp=`wc -l "tmpFile.txt" | grep -oe '[0-9]'`
         echo "countInvAndApp is: " $countInvAndApp
         if [[ $countInvAndApp -ge 2 ]]; then
         for moving in `cat tmpFile.txt`
         do
         echo "moving is: " $moving
         mv -f $moving -t $6/temp
         done
#         echo "Moving " $
#         mv `ls -m1 *.pdf | grep "$invNum"` "$6/temp/"
         cd $6/temp
         gs -dBATCH -dNOPAUSE -sDEVICE=pdfwrite -dSAFER -r600 -sOutputFile=$6/"Invoice"$invNum".pdf" `ls -m1 *.pdf | sort -r`
         rm *.*
         cd "$clear"
        fi
   done
done
fi

cd $6

for file in "$6"/*; do
if [[ -f "invoice.txt" ]]; then
echo "${file}"
file_name="${file##*/}" # get filename with the path
file_path=$(dirname "${file}") # get the path to the file
file_ext="${file##*.}" # get the file extension
file_no_ext="${file_name%.*}" # get filename no extension
file_full=$(basename "$file") # get the filename


#if [[ -f "invoice.txt" ]]; then
if [ "$(ls -A "$6")" ] && [ ${file_ext} == "pdf" ] ; then
########################## GET BARCODE
cd "$1"
rm Thumbs.db
oldest_bc="$(ls *.png | grep -oe '[a-zA-Z]\{1\}[0-9]\{2\}[a-zA-Z]\{2\}[0-9]\{4,25\}' | head -1)" # Get latest file from the barcodes
echo "THE NEXT BARCODE FILE IS:" "${oldest_bc}"
cp "${oldest_bc}"".png" "$2" # back up original barcode image
#gm convert -density 200 "${oldest_bc}"".png" -scale 220% "${oldest_bc}"".miff" #scale the barcode image
gm convert -colorspace Rec709Luma -quality 9 -density 125 "${oldest_bc}"".png" -scale 220% "${oldest_bc}"".miff"
rm "${oldest_bc}"".png"
mv "${oldest_bc}"".miff" "$6" #move the processed barcode image for processing

################################################################
cd "$6"
rm Thumbs.db
echo "PDF Invoice found ...."

echo "PDF split started."

gs -dBATCH -dSAFER -dNOPAUSE -sDEVICE=pdfwrite -sPAPERSIZE=a3 -sOutputFile="${file_no_ext}""_"%003d".pdf" "${file_full}" -c quit

mv "${file_full}" "$4"/"${file_no_ext}"."${today}".pdf #archive the original PDF file

rm Thumbs.db

echo "Start barcode merging and PDF conversion."

firstPage="$(ls -m1 "${file_no_ext}""_"*".pdf" | cut -f 1 -d '.' |  head -1)"

echo "firstPage: " $firstPage

ls -m1 "${file_no_ext}""_"*".pdf" | cut -f 1 -d '.' | sort -n > jpegs.txt
echo "MIFF's :" `cat jpegs.txt`

  for jpeg in `cat jpegs.txt`
   do
   echo "Creating MIFFs."
   echo "MIFF is:" $jpeg

#    gm convert -units PixelsPerInch -density 225 -type Grayscale "${jpeg}"".pdf" "${jpeg}"".jpg"
     gm convert -colorspace Rec709Luma -quality 9 -density 125 "${jpeg}"".pdf" -page A3 "${jpeg}"".miff"
   done

#gm convert -limit pixels 999999999999999 -colorspace Rec709Luma -quality 9 -density 125  "${oldest_bc}"".miff" "${file_no_ext}"".miff" -append -page A4 "${file_no_ext}"".pdf"


gm convert -limit pixels 999999999999999 -colorspace Rec709Luma -quality 9 -density 125  "${oldest_bc}"".miff" "${firstPage}"".miff" -append -page A3 "${firstPage}"".miff"

echo "Convert MIFF to PDF."

gm convert -limit pixels 999999999999999  "${file_no_ext}""_"*".miff" -adjoin -page A4 "${file_no_ext}"".pdf"

#gm convert -limit pixels 999999999999999 -colorspace Rec709Luma -quality 9 -density 125  "${oldest_bc}"".miff" "${file_no_ext}"".miff" -append -page A4 "${file_no_ext}"".pdf"

mv "${file_no_ext}"".pdf" "$5"
#rm *.miff

echo "Cleaning temp files."
rm  "${file_no_ext}""_"*".pdf"

else
echo "PDF Invoice NOT found...."

fi
fi

done

#rm "${file_no_ext}"".tif"
rm invoice.txt
rm jpegs.txt
}

############### Separate processing for Unilever ###############
function check_folder_UNI {

############################ cd Input Invoice folder
cd $3
#echo $3
rm Thumbs.db
#mv *.pdf $6

#Check if invoice.txt exists, should ensure that all files have been saved to the server
if [[ -f "invoice.txt" ]]; then
echo "Begin program execution..."
sleep 7s

mv *.* $6
fi

dateUNI="$(date +%Y%m%d)"
username="$(id -u -n)"
uniDIR="/home/$username/workfolder/UNI/$dateUNI"

if [[ ! -d "${uniDIR}" ]]; then
mkdir "/home/$username/workfolder/UNI/$dateUNI"
fi

cd $6

if [[ -f "invoice.txt" ]]; then

for file in "$6"/*; do
echo "${file}"

file_name="${file##*/}" # get filename with the path
file_path=$(dirname "${file}") # get the path to the file
file_ext="${file##*.}" # get the file extension
file_no_ext="${file_name%.*}" # get filename no extension
file_full=$(basename "$file") # get the filename

#  for invNum in $(cat invoices.txt)
#  do
  if [ "$(ls -A "$6")" ] && [ ${file_ext} == "pdf" ] ; then
########################## GET BARCODE
  cd "$1"
  rm Thumbs.db
  oldest_bc="$(ls *.png | grep -oe '[a-zA-Z]\{1\}[0-9]\{2\}[a-zA-Z]\{2\}[0-9]\{4,25\}' | head -1)" # Get latest file from the barcodes
  echo "THE NEXT BARCODE FILE IS:" "${oldest_bc}"
  cp "${oldest_bc}"".png" "$2" # back up original barcode image
  gm convert -colorspace Rec709Luma -quality 9 -density 125 "${oldest_bc}"".png" -scale 220% "${oldest_bc}"".miff"
  rm "${oldest_bc}"".png"
  mv "${oldest_bc}"".miff" "$6" #move the processed barcode image for processing

################################################################
  cd "$6"
  rm Thumbs.db
  echo "PDF Invoice found ...."

  invNum="$(ls *.pdf |  grep -oe '[0-9]\{4,25\}' | grep -v '[0-9]\{14\}' | head -1)"
  echo "Invoice number is: " "${invNum}"

  rm Thumbs.db
  echo "Start barcode merging and PDF conversion."

   echo "Creating MIFFs."
       gm convert -colorspace Rec709Luma -quality 9 -density 125 "${file_no_ext}"".pdf" -page A3 "${file_no_ext}"".miff"
   echo "MIFF conversion completed."
   mv "${file_full}" "$4"/"${file_no_ext}"."${today}".pdf #archive the original PDF file
 
gm convert -limit pixels 999999999999999 -colorspace Rec709Luma -quality 9 -density 125  "${oldest_bc}"".miff" "${file_no_ext}"".miff" -append -page A3 "${file_no_ext}"".miff"

echo "Convert MIFF to PDF."

gm convert -limit pixels 999999999999999  "${file_no_ext}"".miff" -adjoin -page A4 "${oldest_bc}""_Unilever_""${invNum}"".pdf"
echo "Conversion completed."

echo "Moving ready file."
mv "${oldest_bc}""_Unilever_""${invNum}"".pdf" "$5"

done

else
echo "PDF Invoice NOT found...."
fi
rm *

}

while [ true ]
do
echo "Entering InputInvoices\e5FU"
#check_folder "$DIR_BARC_FU" "$DIR_BARC_FUU" "$DIR_ININV_FU" "$DIR_ININV_FUU" "$DIR_OUTINV_FU" "$DIR_LC_FU"
#    echo "Entering InputInvoices\e5FU"
#    sleep 5s
echo "Entering InputInvoices\e5SE"
#check_folder "$DIR_BARC_SE" "$DIR_BARC_SEU" "$DIR_ININV_SE" "$DIR_ININV_SEU" "$DIR_OUTINV_SE" "$DIR_LC_SE"
#   echo "Entering InputInvoices\e5SE"
#   sleep 5s
echo "Entering InputInvoices\e5IT"
#check_folder "$DIR_BARC_IT" "$DIR_BARC_ITU" "$DIR_ININV_IT" "$DIR_ININV_ITU" "$DIR_OUTINV_IT" "$DIR_LC_IT"
#   echo "Entering InputInvoices\e5IT"
#   sleep 5s
echo "Entering InputInvoices\e5MA"
#check_folder "$DIR_BARC_MA" "$DIR_BARC_MA" "$DIR_ININV_MA" "$DIR_ININV_MAU" "$DIR_OUTINV_MA" "$DIR_LC_MA"
#   echo "Entering InputInvoices\e5MA"
#   sleep 5s
echo "Entering InputInvoices\e5SH"
#check_folder "$DIR_BARC_SH" "$DIR_BARC_SHU" "$DIR_ININV_SH" "$DIR_ININV_SHU" "$DIR_OUTINV_SH" "$DIR_LC_SH"
#   echo "Entering InputInvoices\e5SH"
#   sleep 5s
echo "Entering InputInvoices\e5UT"
#check_folder "$DIR_BARC_UT" "$DIR_BARC_UTU" "$DIR_ININV_UT" "$DIR_ININV_UTU" "$DIR_OUTINV_UT" "$DIR_LC_UT"
#   echo "Entering InputInvoices\e5UT"
#   sleep 5s
echo "Entering InputInvoices\e5OT"
#check_folder "$DIR_BARC_OT" "$DIR_BARC_OTU" "$DIR_ININV_OT" "$DIR_ININV_OTU" "$DIR_OUTINV_OT" "$DIR_LC_OT"
#    echo "Entering InputInvoices\e5OT"
#    sleep 5s
echo "Entering InputInvoices\e5AS"
#check_folder "$DIR_BARC_AS" "$DIR_BARC_ASU" "$DIR_ININV_AS" "$DIR_ININV_ASU" "$DIR_OUTINV_AS" "$DIR_LC_AS"
#    echo "Entering InputInvoices\e5AS"
#    sleep 5s
echo "Entering InputInvoices\UNI"
check_folder_UNI "$DIR_BARC_SH" "$DIR_BARC_SHU" "$DIR_ININV_UNI" "$DIR_ININV_UNIU" "$DIR_OUTINV_UNI" "$DIR_LC_UNI"
#   echo "Entering InputInvoices\UNI"
   sleep 5s

done





#check_folder "$DIR_BARC_FU" "$DIR_BARC_FUU" "$DIR_ININV_FU" "$DIR_ININV_FUU" "$DIR_OUTINV_FU"
#check_folder "$DIR_BARC_SE"
#check_folder "$DIR_BARC_IT"
#check_folder "$DIR_BARC_MA"
#check_folder "$DIR_BARC_SH"
#check_folder "$DIR_BARC_UT"
#check_folder "$DIR_BARC_OT"
#check_folder "$DIR_BARC_AS"



#Echo_string Hello World George
#today=`date '+%d%m%Y_%H%M%S'`;
#echo $today

exit 0

#!/bin/bash

#
# Run this script within the indexing diretory
# It will find all *.vbs files and look for any lines that the provided word
#

if [ $# -ne 1 ]; then
   echo "${0} <word>"
   echo "  builds ./index/<word> for all vbs files"
   exit
fi

TMP_DIR="tmp"
if [ ! -d "${TMP_DIR}" ]; then
    mkdir "${TMP_DIR}"
fi

INDEX_DIR="../index"
if [ ! -d "${INDEX_DIR}" ]; then
    mkdir "${INDEX_DIR}"
fi

INDEX_WORD=$1
FILE_INDEX_FULL_SCAN="${TMP_DIR}/_index.full_scan"
FILE_INDEX="${INDEX_DIR}/${INDEX_WORD}.md"

if [ -f "${FILE_INDEX_FULL_SCAN}" ]; then
    rm "${FILE_INDEX_FULL_SCAN}"
fi
touch "${FILE_INDEX_FULL_SCAN}"

find .. -type f -name "*.vbs" -not -path indexing -print0 | xargs -0 grep --with-filename --only-matching "${INDEX_WORD}" >> "${FILE_INDEX_FULL_SCAN}"
echo "# ${INDEX_WORD}" > "${FILE_INDEX}"
cat ${FILE_INDEX_FULL_SCAN} | sed 's/\.\.\///' | sort | uniq | awk 'BEGIN {FS=OFS=":"} {printf "* [%s](<../%s>)\n", $1, $1}' >> "${FILE_INDEX}"

rm "${FILE_INDEX_FULL_SCAN}"
rmdir "${TMP_DIR}"

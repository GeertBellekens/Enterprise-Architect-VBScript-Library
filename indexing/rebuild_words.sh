#!/bin/bash

#
# Run this script to rebuild the words file, from the contents of the words file
# and the files in the index.
#

WORDS_FILE="words"
INDEX_DIR="../index"

WORDS_FROM_INDEX=""
WORDS_FROM_FILE=""

if [ -d ${INDEX_DIR} ]; then
    WORDS_FROM_INDEX=$(ls -1 ${INDEX_DIR} | sed s/.md//)
fi
if [ -f ${WORDS_FILE} ]; then
    WORDS_FROM_FILE=$(cat ${WORDS_FILE} | tr -d '\r')
fi

WORDS=$(echo ${WORDS_FROM_INDEX} ${WORDS_FROM_FILE} | xargs -n1 echo | sort | uniq)
WORDS=($WORDS)

printf '%s\n' "${WORDS[@]}" > ${WORDS_FILE}

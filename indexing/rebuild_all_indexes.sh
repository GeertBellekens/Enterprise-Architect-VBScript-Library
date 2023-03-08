#!/bin/bash

#
# Run this script to rebuild all indexes based on the words file.
#

WORDS_FILE="words"

echo "Rebuilding all indexes..."
while read -r line
do
  line=$(echo $line | tr -d '\r\n')
  echo "  ${line}..."
  ./create_index_for_word.sh "${line}"
done < "${WORDS_FILE}"
echo "Done"

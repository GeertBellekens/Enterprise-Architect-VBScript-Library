#!/bin/bash

filename="words"
while read -r line
do
  ./create_index_for_word.sh "${line//[$'\r\n']}"
done < "$filename"

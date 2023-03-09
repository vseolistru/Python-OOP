#!/bin/bash

for file in `find -type f -name "*.docx"`
do
   # wc -l $file;
   # stat -c %s $file;
   libreoffice --convert-to pdf $file;
   echo $file;
done

#!/bin/bash

# Read the script IDs from the text file
while IFS= read -r scriptId
do
  # Set the script ID in .clasp.json
  echo "{\"scriptId\":\"$scriptId\", \"rootDir\":\"src\"}" > .clasp.json

  # Push the code to the script
  npx clasp push
done < script-ids.txt

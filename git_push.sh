#!/bin/bash

for i in {1..70}
do
    git add $i.pdf
    git commit -m "Mail"
    git push origin main
done

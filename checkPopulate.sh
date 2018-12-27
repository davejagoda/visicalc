#!/bin/bash

for c in {a..z}
do
    for r in `seq 1 1000`
    do
        echo $c$r
        ./touchCell.py -t visicalc_token.json -n fullyPopulate -w Sheet1 -c $c$r -s 'was empty after first pass'
    done
done

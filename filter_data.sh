#!/bin/bash

iconv -c -f utf-8 -t ascii $1 > $1.tmp
python form.py $1.tmp

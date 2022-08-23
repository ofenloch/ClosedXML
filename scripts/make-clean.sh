#!/bin/bash

find ./ -type d -name "bin" -exec rm -rv {} \;
find ./ -type d -name "obj" -exec rm -rv {} \;
/bin/rm -f *.log

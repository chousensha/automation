#!/usr/bin/env python

import argparse

desc = "Find and print lines from a file that are a certain length"
parser = argparse.ArgumentParser(description=desc)

# length argument
parser.add_argument(
    '-l',
    help = 'Length value',
    dest = 'length',
    type = int,
    required=True
)

# file add_argument
parser.add_argument(
    '-f',
    help = 'Filename',
    dest = 'filename',
    type = str,
    required=True
)

args = parser.parse_args()

with open(args.filename, "r") as f:
	for line in f.readlines():
		# strip the newline character for accurate counting
		if len(line.strip('\n')) == args.length:
			print line

# SnafflerRedEye
Forget to run Snaffler with output flags and trying to parse it? You could run it again with the `-t` flag to get TSV output - or, run the output through this tool to parse it into CSV/JSON format. Even spits out some stats if you're into that sorta thing.

```
SnafflerEyedrops.py -h
usage: SnafflerEyedrops.py [-h] -p PATH [-s] [-oC CSV] [-oJ JSON]

options:
  -h, --help            show this help message and exit
  -p PATH, --path PATH  Path to snaffler output
  -s, --stdout          Write to stdout
  -oC CSV, --csv CSV    Output csv path
  -oJ JSON, --json JSON
                        Output json path
```

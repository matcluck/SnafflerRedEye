import re
import argparse
import csv
import json
from collections import Counter
import xlsxwriter

class Snaffle:
    def replaceNewlines(self):
        self.content = self.content.replace('\\r\\n', '\n').replace('\\r', '\n').replace('\\n', '\n')
        if '\n' in self.content:
            self.multiline = True

    def replaceEscapedSpaces(self):
        self.content = self.content.replace('\\ ', ' ')

    def replaceEqualsChars(self):
        self.content = self.content.replace('=', '\'=')
        self.regex = self.regex.replace('=', '\'=')

    def __init__(self, triageColour, matchRule, readWrite, matchedRegex, size, lastModified, filePath, content):
        self.triageColour = triageColour
        self.matchRule = matchRule
        self.readWrite = readWrite
        self.matchedRegex = matchedRegex
        self.size = size
        self.lastModified = lastModified
        self.filePath = filePath
        self.content = content
        self.multiline = False
        self.replaceNewlines()
        self.replaceEscapedSpaces()
    def __json__(self):
        return {
                'triageColour': self.triageColour,
                'matchRule': self.matchRule,
                'readWrite': self.readWrite,
                'matchedRegex': self.matchedRegex,
                'size': self.size,
                'lastModified': self.lastModified,
                'filePath': self.filePath,
                'content': self.content
                }

    def __str__(self):
        return 'triageColour: %s\nmatchRule: %s\nfilepath: %s\ncontent: %s\n' % (self.triageColour, self.matchRule, self.filePath, self.content)

    def __iter__(self):
        return iter([self.triageColour, self.matchRule, self.readWrite, self.matchedRegex, self.size, self.lastModified, self.filePath, self.content])

def lossParse(snafflerRow, tsv):
    pattern = re.compile(
        r'^\[.*\] \S+ \S+ \[File\]'
        r' '
        r'\{(?P<triageColour>[^<]+?)\}'
        r'\<'
        r'(?P<matchRule>[^|]+?)'
        r'\|'
        r'(?P<readWrite>[^|]+?)'
        r'\|'
        r'(?P<matchedRegex>.*?)(?=\|\d+(\.\d+)?[kmgtp]?B\|\d\d\d\d\-\d\d-\d\d \d\d\:\d\d\:\d\dZ)'
        r'\|'
        r'(?P<size>[^|]+?)'
        r'\|'
        r'(?P<lastModified>[^>]+?)'
        r'\>'
        r'\((?P<filePath>[^)]*?)\)'
        #r' ' - Doing this via regex causes a parsing issue if no content is provided (e.g. when a match is based on extension). Handled when creating snaffleRecord
        r'(?P<content>.*)'
    )

    if (tsv):
        pattern = re.compile (
            r'^\[.*\]\t\S+ \S+\t\[File\]'
            r'\t'
            r'(?P<triageColour>[^\t]+?)'
            r'\t'
            r'(?P<matchRule>[^\t]+?)'
            r'\t'
            r'(?P<readWrite>[^\t]+?)'
            r'\t\t\t'
            r'(?P<matchedRegex>[^\t]+?)'
            r'\t'
            r'(?P<size>[^\t]+?)'
            r'\t'
            r'(?P<lastModified>[^\t]+?)'
            r'\t'
            r'(?P<filePath>[^\t]+?)'
            r'\t'
            r'(?P<content>.*)'
        )

    match = pattern.search(snafflerRow)
    # try parse with content
    try:
        if(csv):
            snaffleRecord = Snaffle(match.group('triageColour'), match.group('matchRule'), match.group('readWrite'), match.group('matchedRegex'), match.group('size'), match.group('lastModified'), match.group('filePath'), match.group('content')[1:])
        if(tsv):
            snaffleRecord = Snaffle(match.group('triageColour'), match.group('matchRule'), match.group('readWrite'), match.group('matchedRegex'), match.group('size'), match.group('lastModified'), match.group('filePath'), match.group('content'))
        return snaffleRecord
    except Exception as e:
        #print(snafflerRow)
        #print(e)
        return None

def tokeniseContent(snaffle, workbook, default_fmt, highlight_fmt):
    text = snaffle.content
    regex = snaffle.matchedRegex

    tokens = []
    last_index = 0
    #print(f"regex: {regex}, text: {text}")
    for match in re.finditer(regex, text):
        #print(match)
        start, end = match.span()
        if start > last_index:
            tokens.append({"text": text[last_index:start], "format": default_fmt})
        tokens.append({"text": text[start:end], "format": highlight_fmt})
        last_index = end
    if last_index < len(text):
        tokens.append({"text": text[last_index:], "format": default_fmt})
    return tokens

def write2CSV(snaffles, outputPath):
    print("Writing snaffles to %s" % outputPath)
    with open(outputPath, mode='w', newline='') as csvFile:
        fieldnames = ['Triage Colour','Match Rule','File Path','Content']
        writer = csv.writer(csvFile)
        writer.writerow(fieldnames)
        writer.writerows(snaffles)

def write2JSON(snaffles, outputPath):
    print("Writing snaffles to %s" % outputPath)
    with open(outputPath, mode='w', newline='') as jsonFile:
        json.dump(snaffles, jsonFile, default=lambda o: o.__json__(), indent=4)

def write2XLSX(snaffles, outputPath, nohighlight):
    print("Writing snaffles to %s" % outputPath)
    
    fields = [
            {'label': 'Triage Colour', 'width': 10},
            {'label': 'Match Rule', 'width': 30},
            {'label': 'Read/Write', 'width': 5},
            {'label': 'Matched Regex', 'width': 30},
            {'label': 'Size', 'width': 10},
            {'label': 'Last Modified', 'width':20},
            {'label': 'File Path', 'width': 40},
            {'label': 'Content', 'width':40}
        ]
    # setup workbook
    workbook = xlsxwriter.Workbook(outputPath)
    worksheet = workbook.add_worksheet()
    
    # write heading row
    fieldnames = [label.get('label', None) for label in fields]
    worksheet.write_row(0,0,fieldnames, workbook.add_format({
        'bold': True,
        'bg_color': '#000000',
        'font_color': '#FFFFFF'
    }))
    dataRow = 1

    # setup filter
    worksheet.autofilter(0, 0, 0, len(fieldnames) - 1)

    # set column width
    for i in range(len(fields)):
        worksheet.set_column(i,i,fields[i]['width'])

    # setup conditional formatting for columns
    formatRed = workbook.add_format({
        'bg_color': '#FF0000',
        'font_color': '#FFFFFF'

    })
    formatGreen = workbook.add_format({
        'bg_color': '#00FF00',
        'font_color': '#000000'

    })
    formatYellow = workbook.add_format({
        'bg_color': '#FFFF00',
        'font_color': '#000000'
    })
    formatBlack = workbook.add_format({
        'bg_color': '#000000',
        'font_color': '#FFFFFF'
    })

    formats = [
        {'colour': 'red', 'format': formatRed},
        {'colour': 'green', 'format': formatGreen},
        {'colour': 'yellow', 'format': formatYellow},
        {'colour': 'black', 'format': formatBlack}
    ]

    default_fmt = workbook.add_format({'font_color': 'black'})
    highlight_fmt = workbook.add_format({'font_color': 'red', 'bold': True})

    for format in formats:
        worksheet.conditional_format('A1:A1048576', {
            'type': 'cell',
            'criteria': '=',
            'value': '"%s"' % format['colour'],
            'format': format['format']
        })

    # write data
    for snaffle in snaffles:
        #print("Processing " + snaffle.matchedRegex)
        snaffle.replaceEqualsChars() # replace = with '= for .xlsx output
        worksheet.write_row(dataRow, 0, snaffle)
        
        if (not nohighlight):
            tokens = tokeniseContent(snaffle, workbook, default_fmt, highlight_fmt)

            if len(tokens) > 1:
                segments = []
                segments.append('```\n') if snaffle.multiline else segments.append('`')

                for token in tokens:
                    segments.append(token["format"])
                    segments.append(token["text"])

                segments.append('\n```') if snaffle.multiline else segments.append('`')

                #print("Writing rich string")
                #print(segments)
                #print([type(x).__name__ for x in segments])

                cell = f"H{dataRow + 1}"
                worksheet.write_rich_string(cell, *segments)

        dataRow = dataRow + 1

    workbook.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-p', '--path', help='Path to Snaffler output', required=True)
    parser.add_argument('-s', '--stdout', action='store_true', help='Write to stdout')
    parser.add_argument('-y', '--tsv', action='store_true', help='Specify that the Snaffler output is TSV formatted')
    parser.add_argument('-oC', '--csv', help='Output csv path')
    parser.add_argument('-oJ', '--json', help='Output json path')
    parser.add_argument('-oX', '--xlsx', help='Output xlsx path')
    parser.add_argument('--no-highlight', action='store_true', help='Disable regex based auto highlighting')

    args = parser.parse_args()

    snaffles = []

    snafflerOutput = open(args.path, 'r', encoding='cp1252')
    for row in snafflerOutput:
        try:
            snaffleRecord = lossParse(row.strip(), args.tsv)
            #print(row)
            #print(snaffleRecord)
            #input()
            if (snaffleRecord != None):
                snaffles.append(snaffleRecord)
                print(snaffleRecord) if args.stdout else ""
        except Exception as e:
            print(e)
            print(row)

    print("Provided log file contained %d snaffles with the following triage counts:\n" % len(snaffles))
    
    triageCounts = Counter(snaffle.triageColour for snaffle in snaffles)

    for key, value in triageCounts.items():
        print("%s: %s" % (key, value))
    
    print("")

    def triageColourToInt(colour):
        if (colour == 'Black') : return 0
        if (colour == 'Red') : return 1
        if (colour == 'Yellow') : return 2
        if (colour == 'Green') : return 3
        return -1      

    sorted_snaffles = sorted(snaffles, key=lambda x: triageColourToInt(x.triageColour))
    
    if args.csv:
        write2CSV(sorted_snaffles, args.csv)

    if args.json:
        write2JSON(sorted_snaffles, args.json)

    if (args.xlsx):
        import xlsxwriter
        write2XLSX(sorted_snaffles, args.xlsx, args.no_highlight)

if __name__ == "__main__":
    main()

import re
import argparse
import csv
import json
from collections import Counter


class Snaffle:
    def __init__(self, triageColour, matchReason, filepath, content):
        self.triageColour = triageColour
        self.matchReason = matchReason
        self.filepath = filepath
        self.content = content
    def __json__(self):
        return {
                'triageColour': self.triageColour,
                'matchReason': self.matchReason,
                'filepath': self.filepath,
                'content': self.content
                }

    def __str__(self):
        return 'triageColour: %s\nmatchReason: %s\nfilepath: %s\ncontent: %s\n' % (self.triageColour, self.matchReason, self.filepath, self.content)

    def __iter__(self):
        return iter([self.triageColour, self.matchReason, self.filepath, self.content])

def lossParse(snafflerRow):
    pattern = re.compile (
        r'^\[.*\] \S+ \S+ \[File\]'
        r' '
        r'\{(?P<triageColour>.*)\}'
        r'\<(?P<matchReason>.*)\>'
        r'\((?P<filepath>[^)]*?)\)'
        r'(?P<content>.*)'
    )

    match = pattern.search(snafflerRow)
    # try parse with content
    try:
        snaffleRecord = Snaffle(match.group('triageColour'), match.group('matchReason'), match.group('filepath'), match.group('content'))
        return snaffleRecord
    except AttributeError:
        # try parse with no content
        try:
            snaffleRecord = Snaffle(match.group('triageColour'), match.group('matchReason'), match.group('filepath'), "")
            return snaffleRecord
        except Exception as e:
            #print(snafflerRow)
            #print(e)
            return None

def write2CSV(snaffles, outputPath):
    print("Writing snaffles to %s" % outputPath)
    with open(outputPath, mode='w', newline='') as csvFile:
        fieldnames = ['Triage Colour','Match Reason','File Path','Content']
        writer = csv.writer(csvFile)
        writer.writerow(fieldnames)
        writer.writerows(snaffles)

def write2JSON(snaffles, outputPath):
    print("Writing snaffles to %s" % outputPath)
    with open(outputPath, mode='w', newline='') as jsonFile:
        json.dump(snaffles, jsonFile, default=lambda o: o.__json__(), indent=4)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-p', '--path', help='Path to snaffler output', required=True)
    parser.add_argument('-s', '--stdout', action='store_true', help='Write to stdout')
    parser.add_argument('-oC', '--csv', help='Output csv path')
    parser.add_argument('-oJ', '--json', help='Output json path')
    args = parser.parse_args()

    snaffles = []

    snafflerOutput = open(args.path, 'r', encoding='cp1252')
    for row in snafflerOutput:
        try:
            snaffleRecord = lossParse(row.strip())
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
    
    if args.csv:
        write2CSV(snaffles, args.csv)

    if args.json:
        write2JSON(snaffles, args.json)

if __name__ == "__main__":
    main()

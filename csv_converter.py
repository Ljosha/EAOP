import pandas as pd
import sys
import argparse
import re

DEBUG = 0

def process_text_file(file_name):
    sentence_pattern = re.compile(r'(?<=[.!?])\s*')

    if file_name == '':
        print('Please specify input file!')
        sys.exit()
        
    # Read the content from the specified text file
    with open(file_name, 'r') as file:
        lines = file.readlines()

    # Process the lines to separate speakers and sentences
    speakers = []
    text4speaker = []
    sentences = []

    current_speaker = None

    for line in lines:
        line = line.strip()
        if DEBUG == 1:
            print(f'Line: {line}')
        
        if line.startswith("Speaker") or line.startswith('Unknown Speaker'):

            if current_speaker is None:
                current_speaker = line
                if DEBUG == 1:
                    print(f'Current Speaker: {current_speaker}')
                text4speaker = []
                
        elif line == '':
            if DEBUG == 1:
                print(f'done with {current_speaker}')
            current_speaker = None
            text4speaker = []
            
        elif line.startswith('Transcribed'):
            break
        
        else:
            if '?' in line or '!' in line or '.' in line:  
                text4speaker = sentence_pattern.split(line)
                text4speaker.pop()
            else:
                text4speaker.extend([line])
            if DEBUG == 1:
                print(f'Text for {current_speaker}: {text4speaker}')
            speakers.extend([current_speaker] * len(text4speaker))
            sentences.extend(text4speaker)
            
    if len(speakers) != len(sentences):
        if DEBUG == 1:
            print(f'Speaker and sentences length dont match')

    # Create a DataFrame
    data = {'Speaker': speakers, 'Sentence': sentences}
    df = pd.DataFrame(data)

    # Write to Excel file
    output_file = file_name.replace('.txt', '_output.xlsx')
    df.to_excel(output_file, index=False)
    print(f"Output written to {output_file}")

if __name__ == "__main__":

    # Create the parser
    parser = argparse.ArgumentParser(description='Your script description here')

    # Add arguments
    parser.add_argument('-f', '--file', type=str, help='Specify the input file with trasncribed interview')
    parser.add_argument('-v', '--verbose', action='store_true', help='Enable verbose mode')

    # Parse the arguments
    args = parser.parse_args()

    input_file = ''
    
    if args.file:
        input_file = args.file
        print(f'Specified file: {args.file}')
    else:
        print('No file specified')

    if args.verbose:
        DEBUG = 1
        print('Verbose mode enabled')
    else:
        print('Verbose mode disabled')
        
    # Process the text file
    process_text_file(input_file)

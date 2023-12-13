# sg-malay-corpus
Data ETL: Development of the Singapore Malay Spoken Words Corpus

## **VBA scripts within this repository:**
1. datacleaning_part1.bas
2. datacleaning_part2.bas

### Goal of VBA scripts 🎯 : 
The VBA scripts within this repository were used to perform ETL of subrip text files sourced from local Singapore Malay television shows and movies into Excel for data cleaning and manipulation, resulting in the creation of a Singapore Malay Spoken Words corpus. 

These VBA scripts resulted in a total of **640,544** words extracted from **227 srt files** across 24 different television shows or movies. 

Here's a sample of the srt file in .txt file format:

<img width="472" alt="Screenshot 2023-12-13 at 11 59 55 AM" src="https://github.com/nurulj11/sg-malay-corpus/assets/145952859/e2512cf3-23c4-4e67-8c3a-98c76177ce4b">


### **datacleaning_part1.bas** was used to achieve the following:
1. Retain words only by removing timestamps and unwanted characters
2. Separate each word into a single cell
3. Compile all words into a single Excel sheet for data manipulation and analysis

The following submacros in **datacleaning_part1.bas** achieved the following: 
A1. Removed timestamps and cells with #NAME? formula name error
A2. Removed punctuation marks in each row 
A3. Converted text to columns, iterated by row, followed by compilation all words in each column into a single column, A. 

_Note: Macros in this script was iterated through each worksheet in the Excel file._


### **datacleaning_part2.bas** was used to achieve the following:
1. Compile all words in each column into a single column, Column A
3. Words from all episodes in each sheet compiled into single sheet

The following submacros in **datacleaning_part2.bas** achieved the following: 
A4. Compiling all columns into a single column on the left i.e. Column A
A5. Copied all words from column A in each sheet (each sheet contains words from a single srt file) into a single sheet

### This is a preview of the final result of the VBA scripts :pencil2:	: 

<img width="1260" alt="Screenshot 2023-12-13 at 12 19 34 PM" src="https://github.com/nurulj11/sg-malay-corpus/assets/145952859/8861d505-bb62-442f-ab91-cc9ae260bd1f">

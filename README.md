# pyxl-examples

This program convert data in the Excel file to text file.

In summary, specifically this code can help you convert data in Excel files into
text files for other processing. The data of each worksheet will be
converted into an independent text file, and there will be corresponding
comments in the text file to identify the name of the worksheet.

## Usage

1. Select and Load the Excel file to be converted.
2. Use the name of the worksheet as part of the text file and generate the
   corresponding text file name.
3. Open the text file and write a comment with the worksheet name.
4. Iterate through each column of the worksheet.
5. Extract the channel label in column 1 and the channel name in column 3.
6. Convert channel labels to channel IDs according to certain rules.
7. Write the channel ID and channel name into a text file.
8. Close the text file.
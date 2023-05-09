# VBA-challenge
Module 2 challenge project to read and parse Stock Ticker through VBA

# Notes
- There is a section in Conditional Formatting requirements that may suggest the need to apply colour backgrounds to the "Percent Change" column, but that is not shown in the provided screenshots. It could be a reference to the % formatting. I chose to follow the screenshots and assume the reference to formatting meant percent style. The **bolded** code below is what I would add for the second conditional row.
If ws.Cells(tickercount + 1, 10).Value > 0 Then
    ws.Cells(tickercount + 1, 10).Style = "40% - Accent3"
    **ws.Cells(tickercount + 1, 11).Style = "40% - Accent3"**
ElseIf ws.Cells(tickercount + 1, 10).Value < 0 Then
    ws.Cells(tickercount + 1, 10).Style = "40% - Accent2"
    **ws.Cells(tickercount + 1, 11).Style = "40% - Accent2"**
End If

# Citations
Subjects not covered in class but researched for this project included:
- Sorting: I used this in several places.
    1. As a failsafe at the start of the code in case the stock info comes unsorted
    2. As a more elegant solution to find the greatest % increase and decrease rather than a brute force loop
    3. As the way to find the largest volume, as well

The sorting with function was learned from the website TrumpExcel.com // https://trumpexcel.com/sort-data-vba/

- Formatting: I used to change the width of the cells and center some of the text for readability

This was leanred from the website AutomateExcel.com // https://www.automateexcel.com/vba/center-text-alignment/

- Accent Styles for background

This was learned by recording a Macro, manually changing the cell background colour to the lighter shades of red/green and then reading the script in VBA for the Macro recorded.


## Instructions
Create a script that loops through all the stocks for one year and outputs the following information:

- The ticker symbol
- Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
- The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
- The total stock volume of the stock. The result should match the following image:

## Moderate solution

Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:

## Hard solution

Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

NOTE
Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

## Other Considerations
Use the sheet alphabetical_testing.xlsx while developing your code. This dataset is smaller and will allow you to test faster. Your code should run on this file in under 3 to 5 minutes.

Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with the click of a button.

## Requirements
Retrieval of Data (20 points)
The script loops through one year of stock data and reads/ stores all of the following values from each row:
- ticker symbol (5 points)
- volume of stock (5 points)
- open price (5 points)
- close price (5 points)

Column Creation (10 points)
On the same worksheet as the raw data, or on a new worksheet all columns were correctly created for:
- ticker symbol (2.5 points)
- total stock volume (2.5 points)
- yearly change ($) (2.5 points)
- percent change (2.5 points)

Conditional Formatting (20 points)
- Conditional formatting is applied correctly and appropriately to the yearly change column (10 points)
- Conditional formatting is applied correctly and appropriately to the percent change column (10 points)

Calculated Values (15 points)
All three of the following values are calculated correctly and displayed in the output:
- Greatest % Increase (5 points)
- Greatest % Decrease (5 points)
- Greatest Total Volume (5 points)

Looping Across Worksheet (20 points)
- The VBA script can run on all sheets successfully.
- GitHub/GitLab Submission (15 points)

All three of the following are uploaded to GitHub/GitLab:
- Screenshots of the results (5 points)
- Separate VBA script files (5 points)
- README file (5 points)

## Grading
This assignment will be evaluated against the requirements and assigned a grade according to the following table:

Grade	Points
A (+/-)	90+
B (+/-)	80–89
C (+/-)	70–79
D (+/-)	60–69
F (+/-)	< 60

## Submission
To submit your Challenge assignment, click Submit, and then provide the URL of your GitHub repository for grading.

IMPORTANT
It is your responsibility to include a note in the README section of your repo specifying code source and its location within your repo. This applies if you have worked with a peer on an assignment, used code in which you did not author or create sourced from a forum such as Stack Overflow, or you received code outside curriculum content from support staff such as an Instructor, TA, Tutor, or Learning Assistant. This will provide visibility to grading staff of your circumstance in order to avoid flagging your work as plagiarized.

References
Data for this dataset was generated by edX Boot Camps LLC, and is intended for educational purposes only.

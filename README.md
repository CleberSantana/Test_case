# Test_case
Software built to read a test case file in .xlsx and show the steps in a "friendly" way.
The file has a specific pattern, then it's possible to automate its reading and writing.

> For using this app it's necessary to install openpyxl and pysimplegui dictionaries.

### Software behavior:
```
 - The interface is loaded 
 - The user must choose a file through file dialog
 - Clicking on "OK" button the sw will check on file all suitable tabs (test cases)
 - A list of all tabs will be displayed
    - According to the test status the color of the line on list is different (green, red, orange)
 - Choosing one item of the list, all related tests will be in memory to navigate
    - According to the test status the color of the line on list is different (green, red, orange)
 - Case a change in "Actual behavior" is performed, this change will be saved in memory
 - Once another tab has been chosen, the file sw save all the infirmations on .xlsx file
 ```
 

# Test_case
Software built to read a test case file in .xlsx and show the steps in a "friendly" way.
the file has a specific pattern, then it's possible to automate its reading.

## Software behavior
```
 - The interface is loaded 
 - The user must choose a file through file dialog
 - Clicking on "OK" button the sw will check on file all suitable tabs (test cases)
 - A list of all tabs will be displayed
 - Choosing one item of the list, all related tests will be in memory to navigate
 - Case a change in "Actual behavior" is performed, this change will be saved in memory
 - Once another tab has been chosen, the file sw save all the infirmations on .xlsx file
 ```

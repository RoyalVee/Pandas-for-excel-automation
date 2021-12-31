# Pandas-for-excel-automation
  Python scripts for excel data automation.
  
## Required software 
  Jupyter notebook
  
  
## Progress Completed 
  The following are completed on the script for section 1:
  - Load data from Marketing Input and Finance Input xlsx
  - identify sheet in the load command
  - Merge all data in all the sheets into one dataframe
  
## To be completed pending project Approval
Part 1:
3 Unpivot the month columns from Jan 21 - Dec 22, so that months will show in column Month, and values in column "Value"
4 The new data set should have columns: Department, Segment, Brand, Market, GLName, Notes, Month, Value (remove 1 2 3 4 5 6 7)
5 save new dataframe into a new xlsx file called Summary.xlsx

Part 2:
1 Load data from the new Summary.xlsx file
2 Split data into 2 dataframes based on the department column. Each department should have its own dataframe
3 Update the Marketing output.xlsx(Data sheet) with the new data. Append below the existing data, leave data before
4 Update the Finance output.xlsx(Data sheet) with the new data. Append below the existing data, leave data before
5 Important that the existing data should not change in the files marked Output, I want to be able to refresh the pivot

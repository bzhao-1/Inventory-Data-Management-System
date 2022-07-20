# Inventory Data Management Tracking System 
# Author: Ben Zhao July 2022

<img width="926" alt="Capture" src="https://user-images.githubusercontent.com/96789119/180080153-cca930da-5931-4eb7-8e84-809b60e3d580.PNG">

How to use: Count items currently in inventory and sort items by category in excel template. See attached screenshot for example. 
Total items currently in inventory should be placed in the column title "Last" for the starting inventory date. Looking at our template, inventory was started on July 7th 2022 with 31 boxes of XS gloves in stockroom. 
When taking items out of stockroom, simply add date to appropriate column and number of items removed from stockroom as seen in template. Run inventoryautomate.py and an updated excel sheet with the number of items currently remaining for each item will be shown. In the result spreadsheet, the column "last" will be deleted. This col is mainly for calculations purposes.
Important Notes: Ensure that sheet names in Python script when reading starting excel file matches your sheet names. Ensure that there is only a single starting excel file to be read. When new shipments arrive, add the number of new items to the number in the "last" column for that item in the template. This will make the calculations for the remaining current number of items accurate on the most recent day.

# Spreadsheet Link updater
Recalculates and updates referenced cells from a spreadsheet. This application was intended to only work with linked cells within an excel spreadsheet in the format solidworks sets up when creating a design table.

## Functions
### Linked Updater and Model Updater
This program is meant to automatically recalculate the reference formulas for each cell from another or the same spreadsheet and saves those values. The excel editing feature in Solidworks does not do this automatically when loading values from a spreadsheet and instead uses the last calculated values. To work around this you'd have to open up the indiviual spreadsheet with excel and save the file to recalculate them. When working on large scale projects in solidworks with a lot of spreadsheet being used to store values this ends up taking a significant amount of time so this program is meant to automate the process.

### Model Updater
Additionally this application can change the formulas from a given spreadsheet inorder to reference a difference set of cells (ex: reference C6 from a sheet instead of B6). In the case of solidworks the is useful if you have multiple variants of your design but the only factor that changes is the overall dimensions. Assumung there is a different set of cells you can reference to account for those changes all that would need to happen is to change which cells your formulas are referencing.

## Using
To use this application on windows, run the exe file. Haven't tested this with other OS but, you can run the script file to use the application as well.

After running it once, a directories.txt and log.txt file will be generated where the script/application file is. Edit the directories file to include all the folder locations where your spreadsheets are. Running the application again (assuming the folder locations exists) will allow you to run to application for it's intended purpose. The log.txt file generated will show the exact files or directories that had issues loading or editing.

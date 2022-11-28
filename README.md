# Navisworks Clash Exporting

## Installation

You can Download the installer [Here](https://github.com/Coolicky/Navisworks.Clash.Exporter/releases/tag/v0.8.0)

##### Installation Location

The plugin should be installed `%AppData%\Autodesk\ApplicationPlugins\Navisworks.Clash.Exporter.bundle`

### Using Directly from Navisworks

Inside the Navisworks Manage You should see an additional tab "Export add-ins 1" with a new button "Export Clashes".

Once clicked it will ask You to specify an Excel file (.xlsx). The application should export a clash report (see details below)

## Excel Export

The Exported Excel file should contain:

- Clash Test Summary
  
  > Summary will contain list of tests with clash count per status (New, Active, Reviewed, Approved, Resolved, Total)

- Clash Tests Details
  
  > `Name` Name of the Clash Test
  > 
  > `Guid` Unique Test Identifier
  > 
  > `No. Clashes` Count of clashes in the test
  > 
  > `Status`
  > 
  > `Last Run` Date & Time of the last time the test was run
  > 
  > `Test Type`
  > 
  > `Tolerance` Tolerance Distance (mm by default, change using optional arguments, see below)

- Clash Groups Details
  
  > `Test Guid` Reference to the Clash Test to which the Group belongs
  > 
  > `Name` Name of the Clash Group
  > 
  > `Guid` Unique Test Identifier
  > 
  > `No. Clashes` Count of clashes in the test
  > 
  > `Approved By`
  > 
  > `Approved Time` (If none will display default Unix Time (01/01/1970))
  > 
  > `Assigned To`
  > 
  > `Center` The Coordinates of the center of the clash
  > 
  > `Created Time`
  > 
  > `Description`
  > 
  > `Distance` The distance of the clash overlap (For Group the Most severe clash) (mm by default, change using optional arguments, see below)
  > 
  > `Status`

- Clash Results Details
  
  > `Test Guid` Reference to the Clash Test to which the Group belongs
  > 
  > `Group Guid` Reference to the Clash Group to which the Test belongs
  > 
  > `Name` Name of the Clash Group
  > 
  > `Guid` Unique Test Identifier
  > 
  > `No. Clashes` Count of clashes in the test
  > 
  > `Approved By`
  > 
  > `Approved Time` (If none will display default Unix Time (01/01/1970))
  > 
  > `Assigned To`
  > 
  > `Center` The Coordinates of the center of the clash
  > 
  > `Created Time`
  > 
  > `Description`
  > 
  > `Distance` The distance of the clash overlap (For Group the Most severe clash) (mm by default, change using optional arguments, see below)
  > 
  > `Status`
  > 
  > `Item 1 Guid` Reference to the First Clashing Item
  > 
  > `Item 2 Guid` Reference to the Second Clashing Item
  > 
  > `Grid Intersection` Name of the closest Grid Intersection to the center of the clash
  > 
  > `Level` Name of the Closest Level to the center of the Clash

- Clashing Elements Details
  
  > `Name` Name of the Clash Group
  > 
  > `Guid` Unique Test Identifier
  > 
  > `ClassName` Class Name (type) of the element
  > 
  > `Model` The name of the source model to which the element belongs
  > 
  > Additional Columns can be added using [Quick Properties](%5BHelp%5D(https://help.autodesk.com/view/NAV/2020/ENU/?guid=GUID-1555C5C2-923B-4342-8120-6BB0EADF45E1)). Each Quick Property will display as additional column for the elements

- Clash Comments
  
  > `Owner Guid` Identifier of the owner (either Clash, Group or Test)
  > 
  > `ID`
  > 
  > `Author`
  > 
  > `Body` The text of the comment itself
  > 
  > `Status`
  > 
  > `Creation Date`

- Historical Summary
  
  > **Only for Automatic Setup**
  > 
  > Historical Summary is almost identical to Summary Page. However it will additionally include a `Date` Column. Previous summaries will be saved in this Page allowing for comparison over time.

The Excel can be used as a Data Source for Power Bi Report and relationship between the tables can be established using the `Guid` values.

## Automation

#### Features

The application will

1. Re-Run all previously set up tests.

2. Group Clashes (if required)

3. Exports Clash Report to Excel (.xlsx)

4. Saves Previous Exports (if required)

5. Saves the Navisworks File

#### Set-up

##### Command Line

You can run the Automation from Command Line. Either create an empty .cmd file or create a new task in Task Scheduler.

Point to the automation executable.`%AppData%\Autodesk\ApplicationPlugins\Navisworks.Clash.Exporter.bundle\Automation\VERSION\Navisworks.Clash.Exporter.Automation.exe`

And provide appropriate arguments

> `-n, --navisworks`
> 
> followed by the path to the Navisworks file (.nwd/.nwf) inside quotation marks
> 
> e.g. `-n "C:\Folder\File.nwf"`
> 
> or `--navisworks "C:\Folder\File.nwf"`

> ``-f, --exportFolder``
> 
> followed by the path the folder where exported Excel file will be saved inside quotation marks
> 
> e.g. `-f "C:\Folder\Export"`
> 
> or `--exportFolder "C:\Folder\Export"`

> `-l, --logLocation`
> 
> followed by the path to the folder where logs will be saved
> 
> e.g. `-l "C:\Folder\Logs"`
> 
> or `--logLocation "C:\Folder\Logs"`

##### Optional Arguments

In addition to required arguments above you can provide further optional arguments

> `--groupBy`
> 
> followed by a grouping. For options see below.
> 
> e.g. `--groupBy Level`

> `--thenBy`
> 
> followed by a grouping. For options see below.
> 
> e.g. `--thenBy GridIntersection`

> Grouping Options:
> 
> > `Level`
> > 
> > `GridIntersection`
> > 
> > `SelectionA`
> > 
> > `SelectionB`
> > 
> > `ModelA`
> > 
> > `ModelB`
> > 
> > `AssignedTo`
> > 
> > `ApprovedBy`
> > 
> > `Status`
> > 
> > `ItemTypeA`
> > 
> > `ItemTypeB`

> `--keepGroups` To keep existing groups

> `--imperial` To export using Imperial Unit System (Will convert distance measurement to feet)

> `--skipRefresh` To skip re-running clash tests

> `--savePrevious` To save previous reports (will copy to "Previous" folder with appropriate time tamp)

> `--skipFileSave` Will skip saving the navisworks file.

Keep all the arguments in a single line. The final result should look similar to the one below: `%AppData%\Autodesk\ApplicationPlugins\Navisworks.Clash.Exporter.bundle\Automation\2022\Navisworks.Clash.Exporter.Automation.exe -n "C:\Folder\File.nwf" -f "C:\Folder\Export" -f "C:\Folder\Export" --groupBy Level --savePrevious`

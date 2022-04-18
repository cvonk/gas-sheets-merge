# Merge Two Google Sheets (Google Apps Script for Sheets)

[![GitHub Discussions](https://img.shields.io/github/discussions/cvonk/gas-sheets-merge)](https://github.com/cvonk/gas-sheets-merge/discussions)
![GitHub tag (latest by date)](https://img.shields.io/github/v/tag/cvonk/gas-sheets-merge)
![GitHub](https://img.shields.io/github/license/cvonk/gas-sheets-merge)


Combines the information from two Google Sheets.  For example, the script can be used to 
combine *people* information exported from a LDAP system with a project assignment list.

## Use

The identifier of the spreadsheet and names of the sheets are found at the end of
the source code. Even if you don't ready anything else, please read through the 
example.

The script [Project Role Allocations](https://github.com/cvonk/gas-sheets-project_role_alloc)
can be used to visualize the resulting data in a pivot table.

## Definitions

The first row of the src/dst sheet should contain labels for each column.
The left-most label in the dst sheet is considered the keyLabel, and the values in that
column are considered *keys*. These keys are used to match rows between the src and dst
sheets.

## Functionality

The function `doMerge()` should be called from the destination sheet.
It walks the src sheet row-by-row.  If the key doesn't already exist in the dst, a new
row is added to the dst.  In the dst sheet, the row with the corresponding key is populated
with values from the src sheet.

## Example

In this example, we start with a source sheet exported from a LDAP system containing the 
usernames, their preferred names and the person type.
The destination sheet already contains some of the usernames imported earlier. These
names were manually supplmented with the projects that they are working on.
The script will add the missing the usernames from the source sheet to the destination
sheet.  Then it writes the preferred names and person type for each user.

### Input

**Source Sheet** (*go-persons*) in *Source Spreadsheet*

| Preferred Name | Username	 | Person Type |
| -------------- | --------- | ----------- |
| Johan          |	jvonk	   | Employee    |
| Sander         |	svonk    |	Employee   |
| Barrie         |	brlevins |	Employee   |
| Tiger          |	tiger    |	Cat        |
| Owen           |	owen     |	Cat        |

**Destination Sheet** (*persons*)

| Username | Person Type |	Role    |	Project 1 |	Project 2 |
| -------- | ----------- | -------- | --------- | --------- |
| jvonk    |	Employee	 | Student	| School    |	Java      |
| cvonk    |	Employee	 | Adult	  | Chores    |	Embedded  |

## Run the Script

  Create a script such as `onopen.gs` that contains an `opOpen()` function to add an
  item to the Google Scripts menu bar.  Remember to fill in your source spreadsheet
  identifier.  

```javascript

    function onOpen_merge() {
      onMerge({srcSpreadsheetId: "YourSpreadSheetId", // eg output from LDAP
               srcSheetName: "go-persons", 
               dstSheetName: "persons"});    
    }
    function onOpen() {
      SpreadsheetApp.getUi()
         .createMenu("YourName")
         .addItem("Merge with go-persons", 'onOpen_merge').addToUi();
    }
```

  Running `onMerge()` updates the columns in the destination sheet that also exist in the
  source sheet.

## What happens

The keyLabel is `Username` since the top-left entry in the destination sheet. 

In this example, the function `doMerge()` will:

 1. Empty the columns in the dst sheet that will be sourced from the src sheet,
    except for the first column that contains the key values;
    
 2. appends rows for which there is no corresponding key value in the dst.

 3. copy the src columns that also exist in the dst sheet.
    
## Output

The *destination sheet* is updated. The roles and project assignments for the
newly imported roles are still blank and need to be filled in by hand.

| Username | Person Type |	Role   | Project 1 | Project 2 |
| -------- | ----------- | ------- | --------- | --------- |
| jvonk    | Employee	   | Student | School    | Java      |
| cvonk    |         	   | Adult   | Chores    | Embedded  |
| svonk    | Employee	   |         |           |           |
| brlevins | Employee    |         |           |           |
| tiger    | Cat         |         |           |           |
| owen     | Cat         |         |           |           |

The empty *Person Type* for Username `cvonk` indicates that that user no longer exists.  One could now remove the persons that are no longer there.

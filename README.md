# Merge Google Sheets (Google Apps Script for Sheets)

Combines the information from two Google Sheets.  In the example it is used to 
combine *people* information exported from a LDAP system, to a project
assignment list.

## Use

The identifier of the spreadsheet and names of the sheets are found at the end of
the source code. Even if you don't ready anything else, please read through the 
example.

The script [Project Role Allocations](https://github.com/cvonk/gas-sheets-projectrolealloc)
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

**Source Sheet**

| Preferred Name| Username	 | Person Type |
| -------------- | --------- | ----------- |
| Johan          |	jvonk	   | Employee    |
| Sander         |	svonk    |	Employee   |
| Barrie         |	brlevins |	Employee   |
| Tiger          |	tiger    |	Cat        |
| Owen           |	owen     |	Cat        |

**Destination Sheet**

| Username | Person Type |	Role    |	Project 1 |	Project 2 |
| -------- | ----------- | -------- | --------- | --------- |
| jvonk    |	Employee	 | Student	| School    |	Java      |
| cvonk    |	Employee	 | Adult	  | Chores    |	Embedded  |

## Run the Script

  Either from the script editor ("Tools > Script Editor), or by adding a
  bit of code to create a menu item, such as 

```javascript
    function onOpen() {
      SpreadsheetApp.getUi()
         .createMenu("CUSTOM")
         .addItem("Merge people", "onMerge").addToUi();
    }
```

  The sheet names are declared at the end of the script.  Remember to fill in
  your source spreadsheet identifier.  

```javascript
    function onMerge() {
      OnMerge.main(srcSpreadsheetId = "<YOUR_SHEET_ID>",
                   srcSheetName = "go-persons", 
                   dstSheetName = "persons");
    }
```

  Running `onMerge()` generates the sheet called *role-alloc-raw* and
  the pivot table *role-alloc*.  If the sheets already exist, it will 
  reuse them.

## What happens

The keyLabel is "Username" since the top-left entry in the destination sheet. In this example, the function `doMerge()` will:

 1. Empty the columns in the dst sheet that will be sourced from the src sheet,
    except for the first column that contains the key values;
    
 2. appends rows for which there is no corresponding key value in the dst.

 3. copy the src columns that also exist in the dst sheet.
    
The empty *Person Type* and *Role* fields for Username *cvonk* indicates that 
that user no longer exists. 

Also, the roles and project assignments for the newly imported roles are still
blank and need to be filled in by hand.

## Output

**Destination Sheet**

| Username | Person Type |	Role   | Project 1 | Project 2 |
| -------- | ----------- | ------- | --------- | --------- |
| jvonk    | Employee	   | Student | School    | Java      |
| cvonk    |         	   |      	 | Chores    | Embedded  |
| brlevins | Employee    |         |           |           |
| tiger    | Cat         |         |           |           |
| owen     | Cat         |         |           |           |

### Filter

One could now remove the persons that are no longer there.

| Username | Person Type |	Role   | Project 1 | Project 2 |
| -------- | ----------- | ------- | --------- | --------- |
| jvonk    | Employee	   | Student | School    | Java      |
| brlevins | Employee    |         |           |           |
| tiger    | Cat         |         |           |           |
| owen     | Cat         |         |           |           |

# Merge Goole Sheets (Google Apps Script for Sheets)

Merges two Google Sheets

## Use

The identifier of the spreadsheet and names of the sheets are found at the end of
the source code.
Even if you don't ready anything else, please read through the example.

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

| Preferred Name	| Username	| Person Type |
| -------------- | -------- | ----------- |
| Johan          |	jvonk	   | Employee    |
| Sander         |	svonk    |	Employee    |
| Barrie         |	brlevins |	Employee    |
| Tiger          |	tiger    |	Cat         |
| Owen           |	owen     |	Cat         |

**Destination Sheet**

| Username	| Person Type |	Role    |	Project 1 |	Project 2 |
| -------- | ----------- | ------- | --------- | --------- |
| jvonk    |	Employee	   | Student	| School    |	Java      |
| cvonk    |	Employee	   | Adult	  | Chores    |	Embedded  |

## Run the Script

## Output

**Destination Sheet**

| Username	| Person Type |	Role    |	Project 1 |	Project 2 |
| -------- | ----------- | ------- | --------- | --------- |
| jvonk    |	Employee	   | Student	| School    |	Java      |
| cvonk    |         	   |      	  | Chores    |	Embedded  |
| brlevins | Employee    |         |           |           |
| tiger    | Cat         |         |           |           |
| owen     | Cat         |         |           |           |


### step-by-step

The keyLabel is "Animal" since the top-left entry in the dst sheet.  In this example, the
function doMerge() will:

 1. empty the columns in the dst sheet
    that will be sourced from the src sheet,
    except for the first column that
    contains the key values;

    | Animal     | Class  | Food   |
    | ---------- | ------ | ------ |
    | moose      |        | plants |
    | deer       |        | plants |


 2. appends rows for which there is no corresponding key value in the dst.

    | Animal     | Class  | Food   |
    | ---------- | ------ | ------ |
    | moose      |        | plants |
    | deer       |        | plants |
    | dophin     |        |        |
    | pufferfish |        |        |

 3. copy the src columns that also exist in the dst sheet.
    
    | Animal     | Class  | Food   |
    | ---------- | ------ | ------ |
    | moose      | mammal | plants |
    | deer       |        | plants |
    | dophin     | mammal |        |
    | pufferfish | fish   |        |


 4. Note: the empty class value for animal "deer" indicates that it (no longer) exists
    in the src.  One could delete or hide the corresponding row.  Also, the Food for
    "dophin" and "pufferfish" is still blank.

# set_goals
Goal settings template in Excel

## Purpose

Personal goal settings tracking tool. Help to set long-term, short-term goals. Define activities to reach those goals. Update goals by changing status. Analyze results, archive records.

## Support and Donations

Enjoying Template? Then consider to buy me a coffee: https://www.paypal.me/Zhbanko

## Goal Setting Template in Excel with UserForm

### Features of the Template:

* Ability to set Short Term, Long Term Goals (Time Bound, Smart, Actionable, etc)
* Ability to log Activities
* Ability to archive records
* Pivot summary table
* Ability to generate PowerPoint slide with a summary of the goal

### Limitations:

* Does not work with Excel OnLine version [Macro-Enabled]
* Limited Fail-safe protection - not trained users may break the file by renaming, deleting columns, etc
* Limited collaboration - better results are expected for personal use
* Pivot Tables must be manually configured

## Build Procedures

It is possible to re-create provided functionality starting from macro-free excel file by following the procedure below:

1. Start from Macro-free Excel Workbook 'Template_Structure.xlsx'
2. Press Alt+F11 to open VBA Project Editor
3. Press CTRL+M to import:

* File Functions.bas
* File Programs.bas
* File PlanningForm.frm
* File invoke2click.cls

4. Open Class Module, double-click on class invoke2click
5. Copy code content from the class invoke2click
6. Paste code content to:

* Sheet 'Report'
* Sheet 'Planning'

7. Save excel workbook as Macro-Enabled file
8. Populate Input fields in the Worksheet 'Summary'

## Test of template

1. Invoke UserForm:

Double click to the row with id 1 on the worksheet 'Planning'

| Result |        Output        |
|--------|:--------------------:|
| Pass   | UserForm will pop-up |
| Fail   |     No user form     |

2. Failsafe check:

Double click to the row 2 on the worksheet 'Report' 

| Result |        Output        |
|--------|:--------------------:|
| Pass   |  Error is displayed  |
| Fail   | UserForm will pop-up |

3. Fields of the UserForm

Populate several records on the worksheet 'Planning', double-click on this row

| Result |                  Output                  |
|--------|:----------------------------------------:|
| Pass   |   Fields of the UserForm are populated   |
| Fail   | Fields of the UserForm are not populated |

4. Save data

Invoke UserForm from the worksheet Planning, populate fields, press button 'Save'

| Result |                          Output                          |
|--------|:--------------------------------------------------------:|
| Pass   | Records are stored in the worksheets Planning and Report |
| Fail   |                Error or data is not saved                |

5. Update data

Invoke UserForm from the worksheet 'Report', by clicking on existing record, change fields, press button 'Save'

| Result |                           Output                           |
|--------|:----------------------------------------------------------:|
| Pass   | Records are updated on both Worksheets Planning and Report |
| Fail   |                 Error or data is not saved                 |

6. Import Picture

Invoke UserForm from the worksheet 'Report', by clicking on existing record, press button 'Import Picture'. Follow prompt to import picture. Press Save button.

| Result |                 Output                |
|--------|:-------------------------------------:|
| Pass   | Picture is visualized in the UserForm |
| Fail   |   Selected picture is not visualized  |

7. Scroll through records

Create and save several records with UserForm. Invoke UserForm and use buttons 'Up' and 'Down' to scroll through records.

| Result |                        Output                        |
|--------|:----------------------------------------------------:|
| Pass   | All records are correctly visualized in the UserForm |
| Fail   |          Records are not properly visualized         |

8. Generate PowerPoint file

Double click on complete record in the Worksheet 'Report'. Press button 'Generate PowerPoint'.

| Result |                 Output                 |
|--------|:--------------------------------------:|
| Pass   | New PowerPoint presentation is created |
| Fail   | Error or Empty PowerPoint Presentation |

9. Archive Record

Double click on complete record in the Worksheet 'Report'. Press button 'Archive'.

| Result |                                                                 Output                                                                |
|--------|:-------------------------------------------------------------------------------------------------------------------------------------:|
| Pass   | Message and Full record is created on the worksheet '.Archive'. Existing records on the worksheets 'Report' and 'Planning are removed |
| Fail   |                                     Records are not copied to worksheet '.Archive' or not removed                                     |

10. Generate PivotTable

Invoke UserForm and press button 'Generate PivotTable'

| Result |                                                                       Output                                                                       |
|--------|:--------------------------------------------------------------------------------------------------------------------------------------------------:|
| Pass   | Pivot table with name 'MyGoalsN' is generated on worksheet 'Pivot'.  Column A of the worksheet 'Pivot' contains numeric value with a Table Id 'N'. |
| Fail   |                                    Error or Pivot Table is not generated. Numeric Id 'N' is not present in Col.A                                   |

## Code saving procedures

This macro workbook contains code. Follow procedure below to export code to the version control:

1. Save code from the worksheets 'Planning' or 'Report' into the Module Class 'invoke2click'
2. Export Form: 'PlanningForm'
3. Export Modules: 'Programs', 'Functions'
4. Export Class: 'invoke2click'

## Issues

Please submit issues here: https://github.com/vzhomeexperiments/set_goals/issues

## Contributions and Support

1. Consider to contribute by creating more features, for example:

* Adding comments to the code
* Adding Fail-safe logic
* Spelling Mistakes
* etc

 2. Fork this repository, create branch and create Pull-Request
 3. Consider supporting this project by:
 
 Enroll to the course on Udemy with referral code: https://www.udemy.com/course/save-your-time-with-excel-userform/?referralCode=0E6A73E1EE79CB01A2E2
 Buy me a coffee: https://www.paypal.me/Zhbanko
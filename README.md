# asctl - Apps Script Control

## General

Helper script to create
- Google Spreadsheet
- Apps Script with all libraries and services attached
- Shared to colleages

## How to execute

![](./docs/asctl-execute-action.png)

## Test Run and authenticate

Go to Spreadsheet > Open Extension > App-Script

Select `FetchResults` and click `Run` then `Review permissions`

![](./docs/9-run-script.png)

Then authenticate with GCP

![](./docs/10-gcp-permissions.png)

Verify

![](./docs/11-execution-log.png)


## Setup Google Spreadsheet Action Button

At this point, you should already see data being injected in the sheet.

Now lets add a trigger button to reload the data.

Insert > Drawing

![](./docs/12-button-drawing.png)

Right click on new button > 3 dots > Assign a script

![](./docs/13-assign-script.png)

Click button and wait for success message to appear `Finished script`

![](./docs/14-final-test.png)
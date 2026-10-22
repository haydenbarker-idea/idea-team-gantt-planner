"""Builds the classic-designer clipboard item for the email-snapshot loop.

Usage: python make-loop.py <workbook file id> <table id> <folder path> > loop.json
"""
import json, sys, uuid

FILE_ID, TABLE_ID, FOLDER = sys.argv[1], sys.argv[2], sys.argv[3]
DRIVE = "b!Wfo6WBsyYk681P6RW4iHyoTq_Lgd58tAp4wYPPTk0KAZxbK0ud9JQ6ciKAnsvR7_"
SCRIPT_ID = "ms-officescript%3A%2F%2Fonedrive_business_itemlink%2F01HQWHCLRBWYKVTVA57NDK6FH6UVRVXPLB"
EXCEL, ONEDRIVE, OUTLOOK = "shared_excelonlinebusiness", "shared_onedriveforbusiness", "shared_office365"
CONN = {
    EXCEL: "shared-excelonlinebu-c7d8d098-fc73-4d25-9ff5-70642d8c9550",
    ONEDRIVE: "shared-onedriveforbu-6fb0b872-2951-4366-9d83-26b5b5bc3cd4",
    OUTLOOK: "shared-office365-23e62306-6fdd-4699-801b-af3179c96cdc",
}
ITEM = "items('Apply_to_each')"
COPY_ID = "@body('Copy_workbook')?['Id']"
RESULT = "body('Prepare_snapshot')?['result']"


def op(api, operation, params, run_after, metadata=None):
    action = {
        "type": "OpenApiConnection",
        "inputs": {
            "host": {"connectionName": api, "operationId": operation, "apiId": f"/providers/Microsoft.PowerApps/apis/{api}"},
            "parameters": params,
            "authentication": "@parameters('$authentication')",
        },
        "runAfter": run_after,
    }
    if metadata:
        action["metadata"] = metadata
    return action


def update_row(status, note, run_after):
    return op(EXCEL, "PatchItem", {
        "source": "me", "drive": DRIVE, "file": FILE_ID, "table": TABLE_ID,
        "idColumn": "Id", "id": f"@{ITEM}?['Id']",
        "item/Status": status, "item/Note": note,
    }, run_after, {FILE_ID: FOLDER, "tableId": TABLE_ID})


ok = ["Succeeded"]
failed = ["Failed", "Skipped", "TimedOut"]
client_email = f"{RESULT}?['clientEmail']"
body = (
    "<p>Here is the customer snapshot of <b>@{" + RESULT + "?['projectName']}</b> as of @{" + RESULT + "?['asOf']}: "
    "a PDF and an Excel copy (values only, this project only).</p>"
    "<p><b>Forward it to:</b> @{if(empty(" + client_email + "), 'the client (no Client Email is set for this project on the Projects tab)', " + client_email + ")}</p>"
    "<p>Remove this note before sending and add your own message.</p>"
    "<p>Sent by the Project Schedules workbook.</p>"
)

BASE = FOLDER.rsplit("/", 1)[0]
LOG = BASE + "/_snapshot-log"
TEMP = BASE + "/_snapshot-temp"

work = {
    # Excel reads can lag writes by a minute or two, so the Status column can't stop a
    # request being picked up twice. A marker file per request Id in OneDrive can.
    "Create_marker": op(ONEDRIVE, "CreateFile", {
        "folderPath": LOG, "name": f"@{{{ITEM}?['Id']}}.done", "body": "@{utcNow()}",
    }, {}),
    "Mark_working": update_row("Working", "", {"Create_marker": ok}),
    "Copy_workbook": op(ONEDRIVE, "CopyDriveFile", {
        "id": FILE_ID,
        "destination": TEMP + "/snapshot-tmp-@{guid()}.xlsx",
        "overwrite": False,
    }, {"Mark_working": ok}, {FILE_ID: FOLDER}),
    "Wait_for_copy": {"type": "Wait", "inputs": {"interval": {"count": 5, "unit": "Second"}}, "runAfter": {"Copy_workbook": ok}},
    "Prepare_snapshot": op(EXCEL, "RunScriptProd", {
        "source": "me", "drive": DRIVE, "file": COPY_ID, "scriptId": SCRIPT_ID,
        "ScriptParameters/tabName": f"@{ITEM}?['Tab']",
    }, {"Wait_for_copy": ok}),
    "Convert_to_PDF": op(ONEDRIVE, "ConvertFile", {"id": COPY_ID, "type": "PDF"}, {"Prepare_snapshot": ok}),
    "Get_Excel_copy": op(ONEDRIVE, "GetFileContent", {"id": COPY_ID, "inferContentType": True}, {"Convert_to_PDF": ok}),
    "Email_requester": op(OUTLOOK, "SendEmailV2", {
        "emailMessage/To": f"@{ITEM}?['Requester']",
        "emailMessage/Subject": "Snapshot ready: @{" + RESULT + "?['projectName']} - forward to @{if(empty(" + client_email + "), 'the client', " + client_email + ")}",
        "emailMessage/Body": body,
        "emailMessage/Attachments": [
            {"Name": "@{" + RESULT + "?['fileBase']}.pdf", "ContentBytes": "@body('Convert_to_PDF')"},
            {"Name": "@{" + RESULT + "?['fileBase']}.xlsx", "ContentBytes": "@body('Get_Excel_copy')"},
        ],
        "emailMessage/Importance": "Normal",
    }, {"Get_Excel_copy": ok}),
    "Mark_sent": update_row("Sent", "@{utcNow()}", {"Email_requester": ok}),
    "Mark_failed": update_row("Failed", "Failed @{utcNow()} - see the flow's run history", {"Email_requester": failed}),
    # Excel Online keeps the copy open for a while after the script; delete it once it lets go.
    "Wait_before_delete": {"type": "Wait", "inputs": {"interval": {"count": 90, "unit": "Second"}},
                           "runAfter": {"Email_requester": ok + failed}},
    "Delete_temp_copy": op(ONEDRIVE, "DeleteFile", {"id": COPY_ID}, {"Wait_before_delete": ok}),
    "Wait_and_retry_delete": {"type": "Wait", "inputs": {"interval": {"count": 4, "unit": "Minute"}},
                              "runAfter": {"Delete_temp_copy": ["Failed"]}},
    "Delete_temp_copy_again": op(ONEDRIVE, "DeleteFile", {"id": COPY_ID}, {"Wait_and_retry_delete": ok}),
}

actions = {
    "Find_marker": op(ONEDRIVE, "GetFileMetadataByPath", {"path": f"{LOG}/@{{{ITEM}?['Id']}}.done"}, {}),
    "Process_request": {"type": "Scope", "actions": work, "runAfter": {"Find_marker": ["Failed"]}},
}

loop = {
    "type": "Foreach",
    "foreach": "@outputs('List_rows_present_in_a_table')?['body/value']",
    "actions": actions,
    "runAfter": {"List_rows_present_in_a_table": ok},
    "runtimeConfiguration": {"concurrency": {"repetitions": 1}},
}

clip = {
    "id": str(uuid.uuid4()),
    "brandColor": "#486991",
    "connectionReferences": {api: {"connection": {"id": f"/providers/Microsoft.PowerApps/apis/{api}/connections/{name}"}} for api, name in CONN.items()},
    "connectorDisplayName": "Control",
    "icon": "data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMzIiIGhlaWdodD0iMzIiIHZlcnNpb249IjEuMSIgdmlld0JveD0iMCAwIDMyIDMyIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPg0KIDxwYXRoIGQ9Im0wIDBoMzJ2MzJoLTMyeiIgZmlsbD0iIzQ4Njk5MSIvPg0KPC9zdmc+",
    "isTrigger": False,
    "operationName": "Apply_to_each",
    "operationDefinition": loop,
}
print(json.dumps(clip))

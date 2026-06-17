---
name: weekly-spbg-report
description: >
  Generate SPBG FW weekly reports by fetching project data from the internal
  API and adding a new worksheet to the SharePoint Excel file. Use when user
  says "weekly report", "file weekly", "do the weekly", "weekly summary".
---
# Weekly SPBG Weekly Report

Fetches project/task data from http://172.21.56.201:5000 API and adds it as a
new sheet (MM_DD) to the SharePoint Excel report via CDP-automated Edge.

## Prerequisites
- Edge browser authenticated to sercomm365.sharepoint.com
- Internal API accessible at http://172.21.56.201:5000
- PowerShell 5.1+ (for CDP WebSocket)

## Parameters
- PROFILE (default: BASW2-2) - team profile
- WEEK_ENDING (default: current Friday) - YYYY-MM-DD format

## SharePoint
- Site: https://sercomm365.sharepoint.com/sites/SPBGFW_SDCInternal
- File ID: 1AA79744-C466-4852-9C4D-385C3099FBB2
- REST: `/GetFileById('${ID}')/$value`
- Sheet naming: MM_DD (e.g., 06_13)

## Step 1: Fetch Data
GET http://172.21.56.201:5000/api/weekly-summary?profile=${PROFILE}&week_date=${WEEK_ENDING}&days=7
Format: Project, Epic, Task Key, Description, Assignee, Status, Progress, Hours

## Step 2: Start Edge with CDP
Start-Process msedge "--remote-debugging-port=9222","--restore-last-session"
New tab: PUT http://127.0.0.1:9222/json/new -> get webSocketDebuggerUrl

## Step 3: Inject JS via CDP WebSocket

### Critical Fix: SheetJS Export
Excel Online page defines XLSX as non-configurable. Replace export before loading:
```js
var r=await fetch('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');
var c=await r.text();
c=c.replace('if(typeof exports!=="undefined")make_xlsx_lib(exports);else if(typeof module!=="undefined"&&module.exports)make_xlsx_lib(module.exports);else if(typeof define==="function"&&define.amd)define("xlsx",function(){if(!XLSX.version)make_xlsx_lib(XLSX);return XLSX});else make_xlsx_lib(XLSX);if(typeof window!=="undefined"&&!window.XLSX)try{window.XLSX=XLSX}catch(e){}','(function(){var __sj={};make_xlsx_lib(__sj);window.__SheetJS__=__sj;})();');
var b=new Blob([c],{type:'text/javascript'});var u=URL.createObjectURL(b);
var s=document.createElement('script');s.src=u;
await new Promise(function(r2,j2){s.onload=r2;s.onerror=j2;document.head.appendChild(s);});
var lib=window.__SheetJS__;
```

### Critical Fix: Form Digest + Checkout
```js
var dr=await fetch(siteUrl+'/_api/contextinfo',{method:'POST',credentials:'same-origin'});
var d=(await dr.text()).match(/<d:FormDigestValue[^>]*>([^<]+)<\/d:FormDigestValue>/)[1];
await fetch(fileApi+'/UndoCheckOut',{method:'POST',headers:{'X-RequestDigest':d},credentials:'same-origin'});
await fetch(dlUrl,{method:'PUT',headers:{'Content-Type':'application/octet-stream','IF-MATCH':'*','X-RequestDigest':d},body:new Uint8Array(outBuf),credentials:'same-origin'});
```

### Critical Fix: WebSocket Cleanup
Always call `$ws.Dispose()` after `$ws.CloseAsync()` - otherwise subsequent connections get ObjectDisposedException.

## Step 4: Full Pipeline (single Runtime.evaluate call)
async function that: Load SheetJS -> download GetFileById -> lib.read() -> aoa_to_sheet() -> book_append_sheet() -> lib.write() -> get digest -> UndoCheckOut -> PUT upload

## Step 5: Verify
Re-download file, check wb.SheetNames.indexOf('MM_DD') >= 0

## CDP in PowerShell
```ps
$ws=New-Object System.Net.WebSockets.ClientWebSocket
$token=New-Object System.Threading.CancellationTokenSource
$ws.ConnectAsync($wsUrl,$token.Token).Wait(5000)
$payload=@{id=1;method="Runtime.evaluate";params=@{expression=$js;awaitPromise=$true;timeout=120000}}|ConvertTo-Json -Depth 10 -Compress
$ws.SendAsync([Text.Encoding]::UTF8.GetBytes($payload),...,$token.Token).Wait(5000)
Start-Sleep -Seconds 50
$b=New-Object byte[] 131072; $r=$ws.ReceiveAsync($b,$token.Token); $r.Wait(30000)
$resp=[Text.Encoding]::UTF8.GetString($b,0,$r.Result.Count)
$ws.CloseAsync([Net.WebSockets.WebSocketCloseStatus]::NormalClosure,"",$token.Token).Wait();$ws.Dispose()
```

## Troubleshooting
| Issue | Fix |
|-------|-----|
| 403 PUT | Get digest from POST /_api/contextinfo |
| 423 Locked | Call UndoCheckOut before upload |
| Can't redefine XLSX | Use the export replace (Fix 1) |
| XLSX.read not func | require.js conflict -> replace entire export block |
| WebSocket disposed | Call Dispose() after CloseAsync |
| 404 API | Use GetFileById, not file path |
| CDP refused | Edge needs --remote-debugging-port flag |

## Full Script
For ready-to-run automation: `scripts/generate-weekly-report.ps1`
Run: `& "...weekly-spbg-report\scripts\generate-weekly-report.ps1" -Profile "BASW2-2"`

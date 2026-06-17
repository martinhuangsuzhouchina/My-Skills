param([string]$Profile="SPBG_FW2-2",[string]$WeekEnding=(Get-Date -Format "yyyy-MM-dd"),[int]$Days=7,[switch]$SkipDataFetch,[string]$DataFile="")
$ErrorActionPreference="Stop"
$SiteUrl="https://sercomm365.sharepoint.com/sites/SPBGFW_SDCInternal"
$FileId="1AA79744-C466-4852-9C4D-385C3099FBB2"
$SheetName=$WeekEnding -replace "^\d{4}-(\d{2})-(\d{2})$",'$1_$2'
$CdpPort=9222;$EdgePath="C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
Write-Host "=== Weekly SPBG Report ===" -ForegroundColor Cyan
Write-Host "Profile: $Profile"
Write-Host "Week Ending: $WeekEnding (sheet: $SheetName)"
if(-not $SkipDataFetch){
    Write-Host "`nStep 1: Fetching data..." -ForegroundColor Yellow
    $apiUrl="http://172.21.56.201:5000/api/weekly-summary?profile=$Profile&week_date=$WeekEnding&days=$Days"
    try{$data=Invoke-RestMethod -Uri $apiUrl -TimeoutSec 30 -ErrorAction Stop
        Write-Host "  Got $($data.results.Count) task records" -ForegroundColor Green
        $rawJson=$data.results|ConvertTo-Json -Compress -Depth 5}
    catch{Write-Host "  Failed: $_" -ForegroundColor Red;exit 1}
}
elseif($DataFile-and(Test-Path $DataFile)){$rawJson=Get-Content $DataFile -Raw}
else{Write-Host "No data source" -ForegroundColor Red;exit 1}
Write-Host "`nStep 2: Building JS pipeline..." -ForegroundColor Yellow$js = @"
(async function(){
    var r=await fetch('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');
    var c=await r.text();
    c=c.replace('if(typeof exports!=="undefined")make_xlsx_lib(exports);else if(typeof module!=="undefined"&&module.exports)make_xlsx_lib(module.exports);else if(typeof define==="function"&&define.amd)define("xlsx",function(){if(!XLSX.version)make_xlsx_lib(XLSX);return XLSX});else make_xlsx_lib(XLSX);if(typeof window!=="undefined"&&!window.XLSX)try{window.XLSX=XLSX}catch(e){}','(function(){var __sj={};make_xlsx_lib(__sj);window.__SheetJS__=__sj;})();');
    var b=new Blob([c],{type:'text/javascript'});var u=URL.createObjectURL(b);
    var s=document.createElement('script');s.src=u;
    await new Promise(function(r2,j2){s.onload=r2;s.onerror=j2;document.head.appendChild(s)});
    var lib=window.__SheetJS__;
    function cs(o){var x={font:{sz:o.sz||10,name:'Calibri',bold:!!o.bold},alignment:{vertical:'top',wrapText:o.wrap!==false,horizontal:o.center?'center':'left'},border:{top:{style:'thin',color:{rgb:o.bc||'D9E2F3'}},bottom:{style:'thin',color:{rgb:o.bc||'D9E2F3'}},left:{style:'thin',color:{rgb:o.bc||'D9E2F3'}},right:{style:'thin',color:{rgb:o.bc||'D9E2F3'}}}};if(o.bg)x.fill={fgColor:{rgb:o.bg},patternType:'solid'};if(o.fc)x.font.color={rgb:o.fc};if(o.nf)x.numFmt=o.nf;return x}
    var hs=cs({sz:11,bold:true,center:true,bg:'1F4E79',fc:'FFFFFF',bc:'8DB4E2'});
    var es=cs({bg:'F2F7FB'});var os=cs({});var ecs=cs({bg:'F2F7FB'});var ocs=cs({});
    var sm={'Closed':cs({center:true,bg:'E2EFDA',bold:true}),'RD Develop':cs({center:true,bg:'DDEBF7',bold:true}),'In Progress':cs({center:true,bg:'DDEBF7',bold:true}),'DEV Leader Assign':cs({center:true,bg:'FCE4D6',bold:true}),'PA Leader Assgin':cs({center:true,bg:'FCE4D6',bold:true}),'Dev/Driver leader Review Code':cs({center:true,bg:'FFF2CC',bold:true}),'RD IT':cs({center:true,bg:'BDD7EE',bold:true}),'Open':cs({center:true,bg:'FCE4D6',bold:true}),'To Do':cs({center:true,bg:'FCE4D6',bold:true}),'Pending':cs({center:true,bg:'FFF2CC',bold:true})};
    var ss=cs({sz:10,bold:true,bg:'1F4E79',fc:'FFFFFF'});var gs=cs({sz:11,bold:true,bg:'002060',fc:'FFFFFF',bc:'8DB4E2'});
    var su='$SiteUrl';var fa=su+"/_api/web/GetFileById('$FileId')";var du=fa+'/\$value';
    var sn='$SheetName';var rd=$rawJson;var mw=1/40;
    var pm={},po=[];rd.forEach(function(r){var p=r.project||'';if(!pm[p]){pm[p]=[];po.push(p)}pm[p].push(r)});
    var rows=[],merges=[],cri=1;
    var hdr=['Project','Epic','Task Key','Task Description','Assignee','Status','Progress','Man-Weeks','Milestones','Summary'];
    po.forEach(function(pn){
        var ts=pm[pn],ph=0,pp=0,dc=0;
        ts.forEach(function(r){ph+=(r.task_hours||0)*mw;pp+=r.progress||0});
        var doneTs=[],remTs=[];
        ts.forEach(function(r){if(r.status==='Closed'||r.status==='Done'||r.status==='Dev/Driver leader Review Code'||r.status==='RD IT'||r.status==='PA Leader Assgin'){doneTs.push(r);dc++}else{remTs.push(r)}});
        var pjS=cri;var sumText=ts.length+' tasks | '+dc+' done / '+(ts.length-dc)+' remain | '+ph.toFixed(2)+' mw | '+(ts.length>0?Math.round(pp/ts.length*10)/10:0)+'% avg';
        var scts=[{lb:'Completed',ts:doneTs,bg:'E2EFDA',fc:'375623'},{lb:'Remaining',ts:remTs,bg:'FCE4D6',fc:'974806'}];
        scts.forEach(function(sec){
            if(sec.ts.length===0)return;
            var secEpics=[];
            rows.push({0:{v:'',s:es},1:{v:'',s:es},2:{v:sec.lb,s:cs({sz:10,bold:true,bg:sec.bg,fc:sec.fc})},3:{v:'',s:es},4:{v:'',s:es},5:{v:'',s:es},6:{v:'',s:es},7:{v:'',s:es},8:{v:'',s:es},9:{v:'',s:es}});
            cri++;var scS=cri;
            sec.ts.forEach(function(r){
                var ev=(cri%2===0),ds=ev?es:os,cs2=ev?ecs:ocs;
                var sts=sm[r.status]||cs2;
                var pv=r.progress!=null?r.progress/100:null;
                rows.push({0:{v:r.project||'',s:ds},1:{v:(r.epic||'').split('\n')[0],s:ds},2:{v:r.task_key||'',s:cs2},3:{v:(r.task||'').substring(0,80),s:ds},4:{v:r.users||'',s:cs2},5:{v:r.status||'',s:sts},6:{v:pv,s:Object.assign({},cs2,{nf:pv!=null?'0%':'@'})},7:{v:ph,s:Object.assign({},cs2,{nf:'0.00'})},8:{v:'',s:cs2},9:{v:'',s:cs2}});
                cri++;secEpics.push((r.epic||'').split('\n')[0]);
            });
            var scE=cri-1;
            if(secEpics.length>=2){var se=scS,ep=secEpics[0];for(var ei=1;ei<secEpics.length;ei++){if(secEpics[ei]!==ep){if(scS+ei-1-se>=1){merges.push({s:{r:se,c:1},e:{r:scS+ei-1,c:1}})};se=scS+ei;ep=secEpics[ei]}}if(scE-se>=1){merges.push({s:{r:se,c:1},e:{r:scE,c:1}})}}
        });
        var pjE=cri-1;
        if(pjE-pjS>=1){merges.push({s:{r:pjS,c:0},e:{r:pjE,c:0}})}
        if(pjE-pjS>=1){merges.push({s:{r:pjS,c:7},e:{r:pjE,c:7}})}
        if(pjE-pjS>=1){merges.push({s:{r:pjS,c:9},e:{r:pjE,c:9}})}
        rows[pjS-1][9]={v:sumText,s:cs({sz:9,fc:'1F4E79'})};
        var ap=ts.length>0?pp/ts.length/100:0;
        rows.push({0:{v:'',s:ss},1:{v:pn+' Total',s:ss},2:{v:ts.length+' tasks',s:ss},3:{v:'',s:ss},4:{v:'',s:ss},5:{v:'',s:ss},6:{v:ap,s:Object.assign({},ss,{nf:ap>0?'0%':'@'})},7:{v:ph,s:Object.assign({},ss,{nf:'0.00'})},8:{v:'',s:ss},9:{v:'',s:ss}});
        cri++;
    });
    var gt=0,gp=0,gc=0,gd=0;
    po.forEach(function(pn){var ts=pm[pn];gc+=ts.length;ts.forEach(function(r){gt+=(r.task_hours||0)*mw;gp+=r.progress||0;if(r.status==='Closed'||r.status==='Done'||r.status==='Dev/Driver leader Review Code'||r.status==='RD IT'||r.status==='PA Leader Assgin')gd++})});
    var ga=gc>0?gp/gc/100:0;
    var gtSummary=gc+' tasks | '+gd+' done / '+(gc-gd)+' remain | '+gt.toFixed(2)+' mw';
    rows.push({0:{v:'Grand Total',s:gs},1:{v:'',s:gs},2:{v:gc+' tasks',s:gs},3:{v:'',s:gs},4:{v:'',s:gs},5:{v:'',s:gs},6:{v:ga,s:Object.assign({},gs,{nf:ga>0?'0%':'@'})},7:{v:gt,s:Object.assign({},gs,{nf:'0.00'})},8:{v:'',s:gs},9:{v:gtSummary,s:gs}});
    var ws={};
    hdr.forEach(function(h,ci){ws[lib.utils.encode_cell({r:0,c:ci})]={v:h,s:hs}});
    rows.forEach(function(row,ri){for(var ci=0;ci<10;ci++){if(row[ci]!=null){ws[lib.utils.encode_cell({r:ri+1,c:ci})]=row[ci]}}});
    ws['!ref']=lib.utils.encode_range({s:{r:0,c:0},e:{r:rows.length,c:9}});
    ws['!cols']=[{wch:18},{wch:30},{wch:14},{wch:50},{wch:18},{wch:18},{wch:10},{wch:10},{wch:40},{wch:55}];
    ws['!freeze']={xSplit:0,ySplit:1,topLeftCell:'A2',activePane:'bottomLeft'};
    ws['!autofilter']={ref:ws['!ref']};
    ws['!merges']=merges;
    var resp=await fetch(du,{credentials:'same-origin'});
    if(!resp.ok)throw new Error('Download failed: '+resp.status);
    var buf=await resp.arrayBuffer();
    var wb=lib.read(new Uint8Array(buf),{type:'array'});
    var idx=wb.SheetNames.indexOf(sn);
    if(idx>=0){wb.SheetNames.splice(idx,1);delete wb.Sheets[sn]}
    lib.utils.book_append_sheet(wb,ws,sn);
    var outBuf=lib.write(wb,{type:'array',bookType:'xlsx'});
    var dr=await fetch(su+'/_api/contextinfo',{method:'POST',credentials:'same-origin'});
    var dx=await dr.text();var digest=dx.match(/<d:FormDigestValue[^>]*>([^<]+)<\/d:FormDigestValue>/)[1];
    await fetch(fa+'/UndoCheckOut',{method:'POST',headers:{'X-RequestDigest':digest},credentials:'same-origin'});
    await new Promise(function(r3){setTimeout(r3,1000)});
    var upResp=await fetch(du,{method:'PUT',headers:{'Content-Type':'application/octet-stream','IF-MATCH':'*','X-RequestDigest':digest},body:new Uint8Array(outBuf),credentials:'same-origin'});
    return JSON.stringify({success:upResp.ok,status:upResp.status,sheets:wb.SheetNames.length,added:sn,tasks:rd.length});
})();
"@
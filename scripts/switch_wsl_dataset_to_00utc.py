#!/usr/bin/env python3
"""Switch only the 13 selected WSL daily products from 06 UTC to 00 UTC.
Preserves extraction files, names, IDs, coordinates and all other products.
Requires openpyxl and the existing assembled Station_Register.xlsx.
"""
from __future__ import annotations
import argparse, calendar, hashlib, io, json, math, os, shutil, sys
from collections import Counter
from datetime import date,datetime,timezone
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from openpyxl import load_workbook

EXPORT_SHA='5c868b12987990b59fb6959f135de9b05070decf413bbbf00ebfd4fbac22314c'
INTERVAL='00:00-24:00 UTC calendar day'
BUILD='2026-11-16.WSL00.1'

def relpath(value):return Path(str(value).replace('\\','/'))

def sha(p):
 h=hashlib.sha256()
 with Path(p).open('rb')as f:
  for b in iter(lambda:f.read(1048576),b''):h.update(b)
 return h.hexdigest()

def table(w,name):
 s=w[name];heads=[c.value for c in s[1]]
 return heads,[dict(zip(heads,row))for row in s.iter_rows(min_row=2,values_only=True) if any(v is not None for v in row)]

def grid(path,daily):
 out={};prev=None;zero=missing=0
 with Path(path).open(encoding='utf-8-sig')as f:
  for line in f:
   if not line.strip():continue
   a=line.rstrip('\r\n').split('\t')
   if len(a)!=(33 if daily else 13):raise ValueError('Wrong field count: '+str(path))
   y=int(a[0]);m=int(a[1])if daily else None;k=(y,m)if daily else(y,)
   if prev is not None and k<=prev:raise ValueError('Duplicate/unordered date rows: '+str(path))
   prev=k;limit=calendar.monthrange(y,m)[1]if daily else 12
   for i,t in enumerate(a[2 if daily else 1:],1):
    t=t.strip();v=None if t.upper()in['','NA','NAN','-9999']else float(t)
    if v is not None and(not math.isfinite(v)or v<0):raise ValueError('Invalid precipitation: '+str(path))
    if i>limit:
     if v is not None:raise ValueError('Numeric padding value: '+str(path))
     continue
    d=date(y,m,i)if daily else date(y,i,1);out[d]=v
    missing+=v is None;zero+=v==0
 if not out:raise ValueError('Empty product: '+str(path))
 good=[d for d,v in out.items()if v is not None]
 return out,dict(valid_observations=len(good),zero_observations=zero,
  first_valid=str(min(good))if good else '',last_valid=str(max(good))if good else '',
  missing_explicit_calendar_cells=missing)

def reconcile(daily,monthly):
 n=ok=0;mx=0
 for d,v in monthly.items():
  if v is None:continue
  values=[daily.get(date(d.year,d.month,i))for i in range(1,calendar.monthrange(d.year,d.month)[1]+1)]
  if any(x is None for x in values):continue
  diff=abs(sum(values)-v);n+=1;ok+=diff<=.05000001;mx=max(mx,diff)
 return dict(complete_daily_months_compared=n,within_0_05mm=ok,outside_0_05mm=n-ok,
  max_absolute_sum_difference_mm=round(mx,9))

def locate_export(root,explicit):
 if explicit:paths=[Path(explicit)]
 else:
  paths=[root/'working/ihgar_station_update/offline_v3_20261109_205739_593/inputs/Network_station_export.zip',
   root/'working/network_station_export/20261104T124610Z_ae7b05/Network_station_export.zip']
  downloads=Path.home()/'Downloads'
  if downloads.exists():paths+=list(downloads.glob('Network_station_export*.zip'))
 for p in paths:
  if p.is_file()and sha(p)==EXPORT_SHA:return p
 raise ValueError('Audited Network_station_export.zip not found. Supply --export PATH; a different source release requires a separate audit.')

def set_fields(sheet,row_number,changes):
 columns={c.value:c.column for c in sheet[1]}
 for key,value in changes.items():
  if key not in columns:raise ValueError('Missing register column: '+key)
  sheet.cell(row_number,columns[key],value)

def main():
 ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--root',type=Path,required=True)
 ap.add_argument('--export',type=Path);ap.add_argument('--check-only',action='store_true');args=ap.parse_args()
 root=args.root.resolve();dataset=root/'03_Daily_dataset';register=dataset/'Station_Register.xlsx'
 evidence=root/'working/dataset_assembly';oldreport=evidence/'Dataset_Assembly_Verification.json'
 reportpath=evidence/'WSL_00UTC_Verification.json'
 prior=json.loads(oldreport.read_text(encoding='utf-8-sig'))
 if sha(register)!=prior['register_sha256']:raise ValueError('Register changed since last verified assembly. Review changes first; no overwrite attempted.')
 export=locate_export(root,args.export);w=load_workbook(register)
 _,stations=table(w,'Stations');_,products=table(w,'Products');_,history=table(w,'History')
 if len(stations)!=435 or len(products)!=778:raise ValueError('Expected the reviewed 435-station, 778-product collection')
 if Counter(p['resolution']for p in products)!=Counter(daily=435,monthly=343):raise ValueError('Product counts changed')
 sids={s['dataset_station_id']for s in stations}
 if len(sids)!=435:raise ValueError('Duplicate station identities')
 paths={};current_hash={}
 for i,p in enumerate(products,1):
  rel=relpath(p['relative_path'])
  if rel.is_absolute()or '..'in rel.parts:raise ValueError('Unsafe register path')
  dest=dataset/rel
  if dest.is_symlink()or not dest.is_file():raise ValueError('Missing/symlink dataset file: '+str(dest))
  if str(rel).lower()in paths:raise ValueError('Duplicate product destination')
  paths[str(rel).lower()]=p
  h=sha(dest)
  if h!=p['sha256']:raise ValueError('Dataset file differs from register: '+str(dest))
  current_hash[str(rel)]=h
  if p['dataset_station_id']not in sids:raise ValueError('Orphan product')
  if i%100==0:print(f'Checked {i}/778 existing dataset hashes',flush=True)
 wsl=[(i+2,p)for i,p in enumerate(products)if p['source_network']=='WSL_LWF'and p['resolution']=='daily']
 if len(wsl)!=13:raise ValueError('Expected 13 WSL daily products')
 station_rows={s['dataset_station_id']:(i+2,s)for i,s in enumerate(stations)}
 plan=[];total=0
 with ZipFile(export)as z:
  for row,p in wsl:
   sid=p['dataset_station_id'];fn=p['filename']
   src=root/'02_Data_extraction/03_WSL_LWF/Daily_00UTC'/fn
   member='wsl_lwf/daily_0utc/'+fn
   expected=hashlib.sha256(z.read(member)).hexdigest()
   if not src.is_file()or src.is_symlink()or sha(src)!=expected:raise ValueError('Extracted WSL 00UTC source does not match audited export: '+str(src))
   # Current content must be either the approved old or new product.
   old=hashlib.sha256(z.read('wsl_lwf/daily_6utc/'+fn)).hexdigest()
   if p['sha256']not in[old,expected]:raise ValueError('Unexpected existing WSL product: '+fn)
   mp=[q for q in products if q['dataset_station_id']==sid and q['resolution']=='monthly']
   if len(mp)!=1:raise ValueError('WSL monthly link not unique: '+sid)
   monthly=mp[0];mrow=products.index(monthly)+2
   monthly_expected=hashlib.sha256(z.read('wsl_lwf/monthly_calendar_utc/'+monthly['filename'])).hexdigest()
   if monthly['sha256']!=monthly_expected:raise ValueError('WSL monthly product differs from audited source export: '+sid)
   mg,_=grid(dataset/relpath(monthly['relative_path']),False);dg,stats=grid(src,True);check=reconcile(dg,mg)
   if check['outside_0_05mm']or not check['complete_daily_months_compared']:
    raise ValueError('WSL 00UTC daily/monthly reconciliation failed: '+fn)
   total+=check['complete_daily_months_compared']
   sn,s=station_rows[sid]
   if s['daily_file']!=p['relative_path']or s['monthly_file']!=monthly['relative_path']:raise ValueError('Station/product filename mismatch')
   plan.append(dict(station_id=sid,filename=fn,source=str(src),destination=str(dataset/relpath(p['relative_path'])),
    old_sha256=p['sha256'],new_sha256=expected,changed=p['sha256']!=expected,
    daily_row=row,monthly_row=mrow,station_row=sn,stats=stats,monthly_check=check))
 if total!=2732:raise ValueError('Complete comparable WSL month count differs from audited 2,732')
 now=datetime.now(timezone.utc).isoformat();stamp=datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ')
 result=dict(build=BUILD,status='CHECK_ONLY'if args.check_only else 'WSL_00UTC_COPY_AND_MONTHLY_AUDIT_PASSED',
  checked_at_utc=now,wsl_daily_files=13,daily_files_total=435,monthly_files_total=343,
  dataset_hashes_checked_before=778,complete_station_months_compared=total,
  monthly_differences_above_0_05mm=0,export_sha256=EXPORT_SHA,export=str(export),
  daily_interval=INTERVAL,source_extraction_modified=False,other_networks_modified=False,
  common_window_analysis_ready=False,
  limitation='Monthly and daily products share an hourly source: arithmetic consistency is not independent measurement validation. WSL hourly timestamp interpretation remains conditional.',
  changes=plan)
 if args.check_only:
  print(json.dumps(result,indent=2));w.close();return 0
 evidence.mkdir(parents=True,exist_ok=True)
 backup=evidence/('WSL_00UTC_Backup_'+stamp+'.zip')
 with ZipFile(backup,'x',ZIP_DEFLATED)as z:
  z.write(register,'Station_Register.xlsx');z.write(oldreport,'Dataset_Assembly_Verification.json')
  if reportpath.exists():z.write(reportpath,'WSL_00UTC_Verification.json')
  for item in plan:z.write(item['destination'],'Daily/'+item['filename'])
  z.writestr('Change_Plan.json',json.dumps(result,indent=2))
 temporary=[];replaced=[];register_replaced=False;summary_replaced=False
 try:
  # Prepare the register and file replacements before committing changes.
  for item in plan:
   sid=item['station_id'];st=item['stats'];mc=item['monthly_check']
   set_fields(w['Products'],item['daily_row'],dict(source_path=item['source'],sha256=item['new_sha256'],
    accumulation_interval=INTERVAL,selection_decision='SELECTED_WSL_00UTC_USER_DECISION',
    destination_verification='VERIFIED_BYTE_IDENTICAL',**st))
   set_fields(w['Products'],item['monthly_row'],dict(**mc,
    monthly_relation='Calendar monthly product paired with selected WSL 00UTC daily product',
    comparison_interpretation='All complete paired calendar months agree within 0.05 mm; shared hourly source, not independent validation'))
   set_fields(w['Stations'],item['station_row'],dict(daily_interval=INTERVAL,
    common_06utc_status='NOT_COMMON_06UTC_SELECTED_00UTC',daily_valid_observations=st['valid_observations'],
    daily_first_valid=st['first_valid'],daily_last_valid=st['last_valid']))
   # History uses existing columns and adds no competing station identifiers.
   hs=w['History'];heads=[c.value for c in hs[1]]
   record=dict(dataset_station_id=sid,decision='WSL product selection: 00UTC',
    reason='User selected native 00-24 daily records to pair with calendar monthly products; original extraction retained',
    start=st['first_valid'],end=st['last_valid'],source=item['source'],
    precipitation_values_changed='Selected product replaced; no individual source observation edited',resolution=now)
   if item['changed']:hs.append([record.get(h)for h in heads])
   if item['changed']:
    target=Path(item['destination']);temp=target.with_name(target.name+'.wsl00_'+stamp+'.tmp')
    temporary.append(temp);shutil.copyfile(item['source'],temp)
    if sha(temp)!=item['new_sha256']:raise ValueError('Prepared replacement hash mismatch')
  w['History'].auto_filter.ref=w['History'].dimensions
  tmpreg=dataset/('Station_Register_WSL00_'+stamp+'.tmp.xlsx');temporary.append(tmpreg)
  w.save(tmpreg);w.close()
  # Read back the full workbook: only intended WSL fields and history may differ.
  oldw=load_workbook(register,read_only=True,data_only=True);neww=load_workbook(tmpreg,read_only=True,data_only=True)
  for name in oldw.sheetnames:
   _,before=table(oldw,name);_,after=table(neww,name)
   if name=='History':
    if after[:len(before)]!=before:raise ValueError('Existing history altered')
    continue
   if len(before)!=len(after):raise ValueError('Workbook row count changed: '+name)
   for a,b in zip(before,after):
    if a.get('network')=='WSL_LWF'and name=='Stations':
     allowed={'daily_interval','common_06utc_status','daily_valid_observations','daily_first_valid','daily_last_valid'}
    elif a.get('source_network')=='WSL_LWF'and name=='Products':
     allowed=({'source_path','sha256','accumulation_interval','selection_decision','destination_verification',*plan[0]['stats']}if a['resolution']=='daily'else{'monthly_relation','comparison_interpretation',*plan[0]['monthly_check']})
    else:allowed=set()
    if any(a.get(k)!=b.get(k)for k in set(a)|set(b)if k not in allowed):raise ValueError('Unintended register change: '+name)
  oldw.close();neww.close()
  for item in plan:
   if not item['changed']:continue
   target=Path(item['destination']);temp=target.with_name(target.name+'.wsl00_'+stamp+'.tmp')
   if sha(target)!=item['old_sha256']:raise ValueError('Destination changed during operation')
   os.replace(temp,target);replaced.append(item)
  for p in products:
   target=dataset/relpath(p['relative_path']);replacement=next((x for x in plan if x['filename']==p['filename']and p['resolution']=='daily'and p['source_network']=='WSL_LWF'),None)
   expected=replacement['new_sha256']if replacement else p['sha256']
   if sha(target)!=expected:raise ValueError('Final dataset hash audit failed: '+str(target))
  os.replace(tmpreg,register);register_replaced=True
  result.update(files_replaced=len(replaced),files_already_selected=13-len(replaced),
   all_dataset_destination_hashes_verified=778,non_wsl_daily_and_all_monthly_files_unchanged=765,
   register=str(register),register_sha256=sha(register),backup=str(backup))
  updated=dict(prior);updated.update(register_sha256=result['register_sha256'],
   status='COPY_AND_REGISTER_VERIFIED_WSL_00UTC_SCIENTIFIC_SCREENING_PENDING',
   latest_product_change=result,latest_verification_utc=now)
  # The original assembly evidence/report is retained in the backup; current report follows current register.
  temp_report=oldreport.with_suffix('.wsl00.tmp');temporary.append(temp_report)
  temp_report.write_text(json.dumps(updated,indent=2),encoding='utf-8');os.replace(temp_report,oldreport);summary_replaced=True
  tmp=reportpath.with_suffix('.tmp');temporary.append(tmp);tmp.write_text(json.dumps(result,indent=2),encoding='utf-8');os.replace(tmp,reportpath)
 except Exception:
  with ZipFile(backup)as z:
   for item in replaced:Path(item['destination']).write_bytes(z.read('Daily/'+item['filename']))
   if register_replaced:register.write_bytes(z.read('Station_Register.xlsx'))
   if summary_replaced:oldreport.write_bytes(z.read('Dataset_Assembly_Verification.json'))
  raise
 finally:
  for p in temporary:
   if p.exists():p.unlink()
 print(json.dumps(result,indent=2));print('Updated register:',register);print('Audit report:',reportpath)
 return 0

if __name__=='__main__':
 try:sys.exit(main())
 except Exception as e:
  print('STOPPED:',e,file=sys.stderr);print('Keep Excel closed and paste the error. No successful update is claimed.',file=sys.stderr);sys.exit(1)

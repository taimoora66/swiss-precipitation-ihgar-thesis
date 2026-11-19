#!/usr/bin/env python3
"""Assemble the audited native station collection, preserving every observation.

Requires Python 3 and openpyxl. No downloads, gap filling, date shifts, QC removal,
renumbering, monthly generation or deletion. This is not common-window validation.
"""
from __future__ import annotations
import argparse, calendar, csv, hashlib, io, json, math, os, shutil, sys
from collections import Counter
from datetime import date, datetime, timezone
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BUILD = '2026-11-16.1'
ORIGINAL_SHA = '8521e488b1361cfb89235fa47894c2a6d3f83ee2d0748bee865a681e4657733e'
SELECT = {'PROPOSE_PRIMARY_DAILY', 'PROPOSE_PRIMARY_PENDING_SPATIAL_METADATA',
          'PROPOSE_PARK_MONTHLY', 'PARK_MONTHLY_IDENTITY_REVIEW'}
# Recovered from Metadata/Site_history in the supplied common-metadata workbooks.
# These are historical PDF-header locations, not current catalogue coordinates.
LOCATIONS = {
 'P_CERE': dict(longitude=8.5476724, latitude=46.3043582, elevation_m=950,
  easting=2685440, northing=1128709, first_year=1991, last_year=1995,
  source_pdf_sha256='8203a7ed92b4de18f142dfa11e9577f2f78fff6e5bcf24970c6b966c3b6d5e9a',
  source_pdf='P_CERE_1991.pdf'),
 'P_ST_ANT': dict(longitude=8.9719480, latitude=46.1583277, elevation_m=215,
  easting=2718440, northing=1113029, first_year=1979, last_year=1987,
  source_pdf_sha256='8f104fbb18ef8e8b0d914172f7e4ef7c08942a71fd802ca52ed43016d5b904bf',
  source_pdf='P_ST_ANT_1979.pdf')}

def digest(p):
 h=hashlib.sha256()
 with Path(p).open('rb') as f:
  for b in iter(lambda:f.read(1048576), b''): h.update(b)
 return h.hexdigest()

def sheets(path):
 w=load_workbook(path,read_only=True,data_only=True)
 try:
  result={}
  for s in w:
   it=s.values; h=next(it)
   if len(h)!=len(set(h)): raise ValueError('Duplicate column names: '+s.title)
   result[s.title]=[dict(zip(h,r)) for r in it if any(v is not None for v in r)]
  return result
 finally: w.close()

def parse_grid(path, daily):
 obs={}; seen=set(); prev=None; zeros=missing=0
 with Path(path).open(encoding='utf-8-sig',newline='') as f:
  for line,fields in enumerate(csv.reader(f,delimiter='\t'),1):
   if not fields: continue
   if len(fields)!=(33 if daily else 13): raise ValueError(f'{path}: line {line}, wrong field count')
   y=int(fields[0]); m=int(fields[1]) if daily else None
   key=(y,m) if daily else (y,)
   if key in seen or prev is not None and key<=prev: raise ValueError(f'{path}: duplicate/unordered date row')
   seen.add(key);prev=key
   count=calendar.monthrange(y,m)[1] if daily else 12
   for j,token in enumerate(fields[2 if daily else 1:],1):
    token=token.strip()
    v=None if token.upper() in ('','NA','NAN','-9999') else float(token)
    if v is not None and (not math.isfinite(v) or v<0): raise ValueError(f'{path}: negative/nonfinite value')
    if j>count:
     if v is not None: raise ValueError(f'{path}: numeric value on nonexistent date')
     continue
    d=date(y,m,j) if daily else date(y,j,1)
    obs[d]=v
    if v is None: missing+=1
    elif v==0: zeros+=1
 if not obs: raise ValueError(f'{path}: empty product')
 valid=[d for d,v in obs.items() if v is not None]
 return obs,dict(valid_observations=len(valid),zero_observations=zeros,
  first_valid=str(min(valid)) if valid else '',last_valid=str(max(valid)) if valid else '',
  missing_explicit_calendar_cells=missing)

def close_number(a,b):
 return a is not None and b is not None and abs(float(a)-float(b))<1e-8

def source_path(root,p):
 p=Path(str(p))
 if not p.is_file(): raise ValueError('Audited source not found: '+str(p))
 if p.is_symlink(): raise ValueError('Symlink input requires explicit review: '+str(p))
 return p

def write_workbook(path,tables):
 w=Workbook();w.remove(w.active)
 for name,rows in tables.items():
  s=w.create_sheet(name)
  cols=list(dict.fromkeys(k for r in rows for k in r)) or ['note']
  s.append(cols)
  for r in rows:
   s.append([r.get(c) for c in cols])
   for cell in s[s.max_row]:
    if cell.data_type=='f': cell.data_type='s'
  s.freeze_panes='A2';s.auto_filter.ref=s.dimensions
  for cell in s[1]:
   cell.font=Font(bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor='244B63')
   cell.alignment=Alignment(wrap_text=True,vertical='top')
  s.row_dimensions[1].height=42
  for i,c in enumerate(cols,1):
   s.column_dimensions[get_column_letter(i)].width=min(48,max(17,len(c)+2))
  for row in s.iter_rows(min_row=2):
   for cell in row:
    if isinstance(cell.value,str) and len(cell.value)>48:
     cell.alignment=Alignment(wrap_text=True,vertical='top')
 w.save(path);w.close()
 # Read-back verifies that the exported register retains all requested data.
 actual=sheets(path)
 for name,rows in tables.items():
  if len(actual[name])!=len(rows): raise ValueError('Workbook read-back row mismatch: '+name)
  for old,new in zip(rows,actual[name]):
   for k,v in old.items():
    if v not in (None,'') and new.get(k)!=v:
     if isinstance(v,float) and isinstance(new.get(k),(float,int)) and math.isclose(v,new[k],rel_tol=1e-12,abs_tol=1e-9): continue
     raise ValueError('Workbook read-back mismatch: '+name+'/'+k)

def main():
 ap=argparse.ArgumentParser(description=__doc__)
 ap.add_argument('--root',type=Path,required=True)
 ap.add_argument('--audit',type=Path)
 ap.add_argument('--summary',type=Path)
 ap.add_argument('--check-only',action='store_true',help='Validate sources and plan; do not create dataset files')
 args=ap.parse_args();root=args.root.resolve()
 audit=args.audit or root/'working/dataset_assembly_audit/Dataset_Register_Audit.xlsx'
 summary_path=args.summary or audit.with_name('Dataset_Assembly_Summary.json')
 evidence=root/'working/dataset_assembly';output=root/'03_Daily_dataset'
 report_path=evidence/'Dataset_Assembly_Verification.json'
 summary=json.loads(summary_path.read_text(encoding='utf-8-sig'));tabs=sheets(audit)
 for name in ['Stations','Proposed_files','Source_audit','History','Review']:
  if name not in tabs: raise ValueError('Missing audit worksheet: '+name)
 if summary.get('blocking_issues')!=0: raise ValueError('Resolve blocking audit issues before assembly')
 if len(tabs['Source_audit'])!=1067 or summary.get('source_products_verified')!=1067:
  raise ValueError('Unexpected audit scope. This script requires the reviewed 1,067-product collection.')
 selected=[dict(r) for r in tabs['Proposed_files'] if r['decision'] in SELECT]
 if Counter(r['time_resolution'] for r in selected)!=Counter(daily=435,monthly=343):
  raise ValueError('Selection differs from reviewed 435 daily / 343 monthly. Review before proceeding.')
 catalog_path=source_path(root,summary['input_ihgar_original'])
 if digest(catalog_path)!=ORIGINAL_SHA: raise ValueError('Original IH-GAR ZIP hash differs from audited original')
 with ZipFile(catalog_path) as z:
  names=[n for n in z.namelist() if n.endswith('/05_Anagrafica_CH_GG.csv')]
  if len(names)!=1: raise ValueError('Original catalogue not uniquely identified')
  text=z.read(names[0]).decode('utf-8-sig');delimiter=csv.Sniffer().sniff(text[:8000]).delimiter
  original={r['nome_file']:r for r in csv.DictReader(io.StringIO(text),delimiter=delimiter)}
 stations={r['station_identity']:dict(r) for r in tabs['Stations']}
 if len(stations)!=len(tabs['Stations']): raise ValueError('Duplicate register identity')
 for r in stations.values():
  if r['network']=='IHGAR':
   old=original[r['provider_id']]
   for k,oldk in [('longitude','Longitudine'),('latitude','Latitudine'),('elevation_m','Elevazione')]:
    if not close_number(r[k],old[oldk]): raise ValueError('IH-GAR location changed: '+r['station_identity'])
   if r['coordinate_basis']!='ORIGINAL_IHGAR_RETAINED': raise ValueError('IH-GAR coordinate basis not retained')
 if sum(r['network']=='IHGAR' for r in stations.values())!=217: raise ValueError('Not all 217 archive entries retained')
 # Match every product to its earlier source audit, then re-read all source files.
 audited={str(r['source_path']):r for r in tabs['Source_audit']}
 if len(audited)!=1067: raise ValueError('Duplicate source paths in audit')
 stats={}
 for i,(p,r) in enumerate(audited.items(),1):
  source=source_path(root,p)
  if r['copy_status']!='VERIFIED_BYTE_IDENTICAL' or r['expected_sha256']!=r['actual_sha256']:
   raise ValueError('Prior audit mismatch: '+p)
  if digest(source)!=r['expected_sha256']: raise ValueError('Source changed since audit: '+p)
  _,st=parse_grid(source,r['time_resolution']=='daily');stats[p]=st
  for k in ['valid_observations','zero_observations']:
   if st[k]!=r[k]: raise ValueError('Observation-count mismatch: '+p)
  if i%100==0: print(f'Revalidated {i}/1067 source products',flush=True)
 # Recover historical locations without replacing source observations or IDs.
 fixes=[]
 for code,loc in LOCATIONS.items():
  r=stations['SUPSI:'+code]
  for k in ['longitude','latitude','elevation_m']:
   if r.get(k) is not None and not close_number(r[k],loc[k]): raise ValueError('Conflicting SUPSI metadata: '+code+'/'+k)
   r[k]=loc[k]
  r['coordinate_system']='EPSG:4326'
  r['coordinate_basis']='HISTORICAL_PDF_HEADER_FROM_SAVED_SITE_HISTORY'
  r['location_source']='SUPSI_excel_common_metadata_v2.zip / stations/'+code+'.xlsx / Site_history'
  r['location_years']=f"{loc['first_year']}-{loc['last_year']} source yearbooks; not installation dates"
  r['source_easting_lv95_m']=loc['easting'];r['source_northing_lv95_m']=loc['northing']
  r['location_source_pdf']=loc['source_pdf'];r['location_source_pdf_sha256']=loc['source_pdf_sha256']
  r['location_verification']='Recovered saved PDF-header metadata; no fresh independent PDF transcription'
  fixes.append(dict(station_identity=r['station_identity'],action='Historical coordinates recovered',
   reason='Current-catalogue fields were empty; saved historical Site_history already contained location',**loc))
  r['common_06utc_status']='PENDING_PROVIDER_INTERVAL_DOCUMENTATION'
 selected_ids={r['station_identity'] for r in selected if r['time_resolution']=='daily'}
 if len(selected_ids)!=435: raise ValueError('Daily identity count differs from 435')
 products=[];destinations=set();daily_by_id={}
 for r in selected:
  sid=r['station_identity']
  if sid not in stations or sid not in selected_ids: raise ValueError('Unlinked product identity: '+sid)
  fn=str(r['filename'])
  if Path(fn).name!=fn or '/' in fn or '\\' in fn: raise ValueError('Unsafe filename: '+fn)
  sub='Daily' if r['time_resolution']=='daily' else 'Monthly'
  rel=Path(sub)/fn;target=output/rel
  if str(rel).lower() in destinations: raise ValueError('Case-insensitive destination collision: '+str(rel))
  destinations.add(str(rel).lower())
  if r['source_sha256']!=audited[str(r['source_path'])]['expected_sha256']: raise ValueError('Proposal hash differs from source audit')
  if target.exists() and (not target.is_file() or target.is_symlink() or digest(target)!=r['source_sha256']):
   raise ValueError('Existing destination conflicts; no overwrite permitted: '+str(target))
  st=stats[str(r['source_path'])]
  if st['valid_observations']!=r['valid_observations']: raise ValueError('Proposed observation-count mismatch')
  product=dict(dataset_station_id=sid,source_network=r['network'],provider_station_id=r['provider_id'],
   filename=fn,relative_path=str(rel),resolution=r['time_resolution'],units='mm',
   accumulation_interval=r['accumulation_interval'],monthly_relation=r.get('monthly_relation'),
   selection_decision=r['decision'],source_path=r['source_path'],sha256=r['source_sha256'],**st)
  if r['time_resolution']=='daily':
   if sid in daily_by_id: raise ValueError('More than one primary daily representation: '+sid)
   daily_by_id[sid]=r
  products.append(product)
 # Existing unknown files inside the two product folders must not be silently mixed.
 for sub in ['Daily','Monthly']:
  folder=output/sub
  if folder.is_symlink(): raise ValueError('Dataset product folder is a symlink')
  if folder.exists():
   for p in folder.rglob('*'):
    if p.is_file() and str(p.relative_to(output)).lower() not in destinations:
     raise ValueError('Unexpected file in destination product folder; review first: '+str(p))
 # Monthly sums are diagnostics, not an instruction to adjust provider totals.
 monthly_done=0
 for p in products:
  if p['resolution']!='monthly': continue
  dr=daily_by_id[p['dataset_station_id']]
  dg,_=parse_grid(dr['source_path'],True);mg,_=parse_grid(p['source_path'],False)
  n=agree=0;largest=None
  for d,v in mg.items():
   if v is None: continue
   vals=[dg.get(date(d.year,d.month,j)) for j in range(1,calendar.monthrange(d.year,d.month)[1]+1)]
   if any(x is None for x in vals): continue
   delta=sum(vals)-v;n+=1;agree+=abs(delta)<=0.05000001
   largest=abs(delta) if largest is None else max(largest,abs(delta))
  p.update(complete_daily_months_compared=n,within_0_05mm=agree,
   outside_0_05mm=n-agree,max_absolute_sum_difference_mm=largest,
   comparison_interpretation='Diagnostic only: provider monthly and selected daily intervals/source histories may differ; no adjustment')
  monthly_done+=1
  if monthly_done%50==0: print(f'Checked {monthly_done}/343 monthly-to-daily diagnostics',flush=True)
 included=[];not_included=[]
 for sid,r in stations.items():
  if sid not in selected_ids:
   not_included.append(dict(dataset_station_id=sid,network=r['network'],provider_station_id=r['provider_id'],
    station_name=r['station_name'],reason=r['identity_status'],status=r['collection_status']))
   continue
  row=dict(r);row['dataset_station_id']=row.pop('station_identity')
  # Null project IDs are legitimate legacy records, never invent new numbers.
  row['legacy_id_status']='Existing project ID' if row.get('project_id') else 'Legacy archive ID retained; no project number invented'
  for k in ['longitude','latitude','elevation_m','provider_longitude','provider_latitude','provider_elevation_m']:
   if row.get(k) is not None: row[k]=float(row[k])
  for k in ['longitude','latitude','elevation_m']:
   if row.get(k) is None or not math.isfinite(row[k]): raise ValueError('Missing/nonfinite selected station metadata: '+sid+'/'+k)
  if not -180<=row['longitude']<=180 or not -90<=row['latitude']<=90:
   raise ValueError('Selected location outside geographic coordinate bounds: '+sid)
  chosen=[p for p in products if p['dataset_station_id']==sid]
  row['daily_file']=next(p['relative_path'] for p in chosen if p['resolution']=='daily')
  monthly=[p['relative_path'] for p in chosen if p['resolution']=='monthly']
  row['monthly_file']=monthly[0] if monthly else None
  row['analysis_eligibility']='Collection retained; analysis-specific interval, QC and physical-site screening pending'
  if row.get('common_06utc_status')=='NOT_SPATIALLY_ELIGIBLE': raise ValueError('Unresolved spatial metadata in selected identity')
  for k in ['collection_status','daily_source','primary_daily_filename','parked_monthly_filename']:
   row.pop(k,None)
  included.append(row)
 if len(not_included)!=10: raise ValueError('Excluded-identity scope changed')
 nested=[str(p) for p in output.iterdir() if p.is_dir() and p.name not in ['Daily','Monthly']] if output.is_dir() else []
 checks=[dict(check='Native collection scope',result='435 daily + 343 monthly; not a harmonized common-window release'),
  dict(check='Revalidated source integrity and structural validity',result='1067/1067 matching hashes; valid calendar and nonnegative finite values'),
  dict(check='IH-GAR locations',result='217/217 match original archive catalogue'),
  dict(check='Historical SUPSI coordinates',result='2 recovered from saved source metadata; underlying rainfall unchanged'),
  dict(check='Spatial/temporal readiness',result='Scientific eligibility remains conditional; do not treat all native days as one UTC window'),
  dict(check='Older nested dataset folders',result='Retained untouched: '+('; '.join(nested) or 'none'))]
 report=dict(build=BUILD,mode='CHECK_ONLY' if args.check_only else 'ASSEMBLED_NATIVE_COLLECTION',
  checked_at_utc=datetime.now(timezone.utc).isoformat(),sources_revalidated=1067,
  daily_files=435,monthly_files=343,included_register_rows=435,not_included_rows=10,
  original_ihgar_locations_verified=217,historical_supsi_locations_recovered=fixes,
  nested_folders_not_modified=nested,source_observations_modified=False,
  common_window_analysis_ready=False,audit_workbook_sha256=digest(audit),audit_summary_sha256=digest(summary_path),
  audit_evidence_retained=True,output=str(output),checks=checks)
 evidence.mkdir(parents=True,exist_ok=True)
 if args.check_only:
  print(json.dumps(report,indent=2));return 0
 # Preserve evidence before assembly. Never delete audit records or source datasets.
 stamp=datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ')
 with ZipFile(evidence/('Assembly_Evidence_'+stamp+'.zip'),'x',ZIP_DEFLATED) as z:
  z.write(audit,'Dataset_Register_Audit.xlsx');z.write(summary_path,'Dataset_Assembly_Summary.json')
  z.writestr('Recovered_SUPSI_Locations.json',json.dumps(fixes,indent=2))
  for p in [output/'Station_Register.xlsx',report_path]:
   if p.is_file(): z.write(p,'previous/'+p.name)
 copied=reused=0
 for p in products:
  dest=output/p['relative_path'];dest.parent.mkdir(parents=True,exist_ok=True)
  if dest.exists(): reused+=1
  else:
   # Exclusive creation refuses an overwrite even if another process creates a file.
   with Path(p['source_path']).open('rb') as src,dest.open('xb') as dst: shutil.copyfileobj(src,dst)
   copied+=1
  if digest(dest)!=p['sha256']: raise ValueError('Destination hash mismatch: '+str(dest))
  p['destination_verification']='VERIFIED_BYTE_IDENTICAL'
  if (copied+reused)%100==0: print(f'Copied/reused and verified {copied+reused}/778 selected files',flush=True)
 # Recheck sources after copying, guarding changes during the operation.
 for p in products:
  if digest(p['source_path'])!=p['sha256']: raise ValueError('Source changed during assembly: '+p['source_path'])
 history=[dict(r) for r in tabs['History']]
 for r in history:
  sid='IHGAR:'+str(r.get('id'))
  if sid not in selected_ids: raise ValueError('Unlinked IH-GAR history record: '+sid)
  r['dataset_station_id']=sid
 for fix in fixes:
  history.append(dict(dataset_station_id=fix['station_identity'],decision='Metadata correction only',
   reason=fix['reason'],start=str(fix['first_year']),end=str(fix['last_year']),
   source=fix['source_pdf'],precipitation_values_changed=False))
 # Record unresolved reviews in the same workbook, avoiding silent disappearance.
 for r in tabs['Review']:
  history.append(dict(dataset_station_id=r['station_identity'],decision=r['severity'],
   reason=r['reason'],source=r['subject'],resolution=('Historical metadata recovered; see Stations' if r['station_identity'] in ['SUPSI:P_CERE','SUPSI:P_ST_ANT'] and r['severity']=='REVIEW' else r['resolution'])))
 tables={'Stations':included,'Products':products,'History':history,'Not_included':not_included}
 temp=output/('Station_Register_'+stamp+'.tmp.xlsx')
 try:
  write_workbook(temp,tables)
  os.replace(temp,output/'Station_Register.xlsx')
 finally:
  if temp.exists(): temp.unlink()
 report.update(new_files_copied=copied,identical_files_reused=reused,destination_hashes_verified=len(products),
  register=str(output/'Station_Register.xlsx'),register_sha256=digest(output/'Station_Register.xlsx'),
  status='COPY_AND_REGISTER_VERIFIED_SCIENTIFIC_SCREENING_PENDING')
 tmp=report_path.with_suffix('.tmp');tmp.write_text(json.dumps(report,indent=2),encoding='utf-8');os.replace(tmp,report_path)
 print(json.dumps(report,indent=2));print('Register:',output/'Station_Register.xlsx');print('Verification:',report_path)
 return 0

if __name__=='__main__':
 try: sys.exit(main())
 except Exception as exc:
  print('STOPPED:',exc,file=sys.stderr)
  print('No source data were intentionally modified. Any files already copied must be rechecked by rerunning; no successful release is claimed.',file=sys.stderr)
  sys.exit(1)

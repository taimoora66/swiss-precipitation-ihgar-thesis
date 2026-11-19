#!/usr/bin/env python3
"""Read-only source audit and proposed dataset register. No station data are moved or changed.
Creates only working/dataset_assembly_audit/Dataset_Register_Audit.xlsx and a JSON summary.
Candidates are a station collection, not an automatically homogeneous common-window dataset.
Requires openpyxl. Current selection: archive-first, MeteoSwiss 06 UTC, WSL 06 UTC, provider-native SUPSI/Agrometeo.
"""
from __future__ import annotations
import argparse, calendar, csv, gzip, hashlib, io, json, math, os, re, shutil, sys
from collections import Counter, defaultdict, OrderedDict
from datetime import date, datetime, timezone
from pathlib import Path
from zipfile import ZipFile

EXPORT_SHA='5c868b12987990b59fb6959f135de9b05070decf413bbbf00ebfd4fbac22314c'
ORIGINAL_SHA='8521e488b1361cfb89235fa47894c2a6d3f83ee2d0748bee865a681e4657733e'
FOLDERS={'IHGAR':'01_IHGAR','MeteoSwiss':'02_MeteoSwiss','WSL_LWF':'03_WSL_LWF','SUPSI':'04_SUPSI','Agrometeo':'05_Agrometeo'}
DEST={'meteoswiss/daily_6UTC':('MeteoSwiss','Daily_06UTC'),'meteoswiss/daily_0UTC':('MeteoSwiss','Daily_00UTC'),'meteoswiss/monthly':('MeteoSwiss','Monthly'),
'agrometeo/daily':('Agrometeo','Daily'),'agrometeo/monthly':('Agrometeo','Monthly'),'supsi/daily':('SUPSI','Daily'),'supsi/monthly':('SUPSI','Monthly'),
'wsl_lwf/daily_6utc':('WSL_LWF','Daily_06UTC'),'wsl_lwf/daily_0utc':('WSL_LWF','Daily_00UTC'),'wsl_lwf/monthly_calendar_utc':('WSL_LWF','Monthly')}
NAME=re.compile(r'^PCP[DM]_CH_\d{3}_(?:[A-Z]{2}|00)_.+_(17|92|93|94)_\d{4}$')

def H(b):return hashlib.sha256(b).hexdigest()
def filehash(p):
 h=hashlib.sha256()
 with p.open('rb') as f:
  for b in iter(lambda:f.read(1048576),b''):h.update(b)
 return h.hexdigest()
def rows_csv(b,delimiter=','):return list(csv.DictReader(io.StringIO(b.decode('utf-8-sig')),delimiter=delimiter))
def rows_sheet(b,sheet):
 from openpyxl import load_workbook
 w=load_workbook(io.BytesIO(b),read_only=True,data_only=True)
 try:
  it=w[sheet].iter_rows(values_only=True);keys=next(it);return [dict(zip(keys,r)) for r in it]
 finally:w.close()
def net(s):return 'WSL_LWF' if s=='WSL' else s
def provider(n,s):return re.sub(r'^AGROMETEO_','',str(s)) if n=='Agrometeo' else str(s)

def num(token):
 s=str(token).strip()
 if s.upper() in ('','NA','NAN','-9999'):return None
 v=float(s)
 if not math.isfinite(v):raise ValueError('Nonfinite observation token')
 if v<0:raise ValueError(f'Negative precipitation token: {s}')
 return v
def same(a,b):return a is None and b is None or a is not None and b is not None and abs(a-b)<=0.000001

def grid(b,daily=True):
 result={};valid=zero=missing=padding=0;first=last=None
 for line_no,line in enumerate(b.decode('utf-8-sig').splitlines(),1):
  if not line.strip():continue
  fields=line.split('\t')
  if len(fields)!=(33 if daily else 13):raise ValueError(f'Line {line_no}: expected {33 if daily else 13} tab-separated fields, got {len(fields)}')
  year=int(fields[0]);month=int(fields[1]) if daily else None
  key=(year,month) if daily else year
  if key in result:raise ValueError(f'Duplicate date row: {key}')
  if daily and not 1<=month<=12:raise ValueError('Invalid month')
  vals=[num(t) for t in fields[2 if daily else 1:]]
  end=calendar.monthrange(year,month)[1] if daily else 12
  for j,v in enumerate(vals):
   if j>=end:
    padding+=1
    if v is not None:raise ValueError(f'Numeric observation on nonexistent date: {year}-{month}-{j+1}')
    continue
   when=date(year,month,j+1) if daily else date(year,j+1,1)
   if v is None:missing+=1
   else:
    valid+=1;zero+=v==0;first=when if first is None or when<first else first;last=when if last is None or when>last else last
  result[key]=vals
 if list(result)!=sorted(result):raise ValueError('Date rows are not chronological')
 if not result:raise ValueError('No data rows')
 return result,dict(rows=len(result),valid_observations=valid,zero_observations=zero,missing_calendar_cells=missing,calendar_padding_cells=padding,first_valid=str(first or ''),last_valid=str(last or ''))
def day(g,dt):
 v=g.get((dt.year,dt.month));return v[dt.day-1] if v is not None else None

class Source:
 def __init__(self,p):
  self.path=Path(p);self.zip=ZipFile(p) if self.path.is_file() else None
  if self.zip:self.index={x.replace('\\','/'):x for x in self.zip.namelist() if not x.endswith('/')}
  else:self.index={str(x.relative_to(self.path)).replace('\\','/'):x for x in self.path.rglob('*') if x.is_file()}
 def names(self):return list(self.index)
 def find(self,suffix):
  matches=[n for n in self.index if n==suffix or n.endswith('/'+suffix)]
  if len(matches)!=1:raise ValueError(f'{self.path}: expected one {suffix}, found {len(matches)}')
  return matches[0]
 def read(self,name):return self.zip.read(self.index[name]) if self.zip else self.index[name].read_bytes()
 def suffix(self,name):return self.read(self.find(name))
 def close(self):
  if self.zip:self.zip.close()

def locate_export(root,downloads,explicit):
 if explicit:
  p=Path(explicit)
  if not p.is_file() or filehash(p)!=EXPORT_SHA:raise ValueError('--export is not the previously audited numbered export (v1.1.1). Review a different export separately.')
  return p
 candidates=[root/'working/ihgar_station_update/offline_v3_20261109_205739_593/inputs/Network_station_export.zip',root/'working/network_station_export/20261104T124610Z_ae7b05/Network_station_export.zip']
 for base in [downloads,root/'working/network_station_export',root/'working/ihgar_station_update']:
  if base.is_dir():candidates.extend(base.glob('Network_station_export*.zip') if base==downloads else base.rglob('Network_station_export*.zip'))
 for p in dict.fromkeys(candidates):
  if p.is_file() and filehash(p)==EXPORT_SHA:return p
 raise ValueError('Audited Network_station_export.zip not found. Use --export PATH; expected SHA-256 '+EXPORT_SHA)
def locate_package(candidates,check,explicit=None):
 if explicit:candidates=[Path(explicit)]
 for p in dict.fromkeys(candidates):
  if not p.exists():continue
  try:
   src=Source(p)
   if check(src):return src
   src.close()
  except Exception:continue
 raise ValueError('Required evidence package not found or not recognized: '+('explicit input '+str(explicit) if explicit else ' | '.join(map(str,candidates))))


def read_local_rows(path,sheet):return rows_sheet(path.read_bytes(),sheet)
def populated(v):return v is not None and str(v).strip()!=''
def identical_metadata(a,b):
 if not populated(a) and not populated(b):return True
 try:return math.isclose(float(a),float(b),abs_tol=1e-7,rel_tol=0)
 except (ValueError,TypeError):return str(a or '')==str(b or '')

def write_book(path,sheets):
 from openpyxl import Workbook,load_workbook
 from openpyxl.styles import Font,PatternFill,Alignment
 from openpyxl.utils import get_column_letter
 b=Workbook();b.remove(b.active)
 for name,rows,cols in sheets:
  s=b.create_sheet(name);s.append(cols)
  for r in rows:s.append([r.get(k,'') for k in cols])
  for row in s:
   for cell in row:
    if cell.data_type=='f':cell.data_type='s'
  for cell in s[1]:cell.font=Font(bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor='234B61');cell.alignment=Alignment(wrap_text=True)
  s.freeze_panes='A2';s.auto_filter.ref=s.dimensions;s.row_dimensions[1].height=32
  for i,k in enumerate(cols,1):s.column_dimensions[get_column_letter(i)].width=62 if any(x in k for x in ('source','path','reason','details','destination','basis')) else 26
 b.save(path)
 w=load_workbook(path,read_only=True,data_only=False)
 try:
  for name,rows,cols in sheets:assert w[name].max_row==len(rows)+1 and w[name].max_column==len(cols)
 finally:w.close()

def main():
 ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--root',required=True,type=Path);ap.add_argument('--export',type=Path);ap.add_argument('--ihgar-release',type=Path);ap.add_argument('--ihgar-original',type=Path);ap.add_argument('--stage2',type=Path);ap.add_argument('--metadata-root',type=Path);ap.add_argument('--audit-output',type=Path)
 a=ap.parse_args();root=a.root.resolve();stage=a.stage2 or root/'02_Data_extraction';metadata_root=a.metadata_root or root/'01_Data_downloaded';audit=a.audit_output or root/'working/dataset_assembly_audit';destination=root/'03_Daily_dataset';downloads=Path.home()/'Downloads';stamp=datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ')
 import openpyxl
 if not stage.is_dir():raise ValueError('Second-stage extraction folder not found: '+str(stage))
 # Audit directory must never be a station product folder.
 if any(audit.resolve().is_relative_to(p.resolve()) for p in (stage,destination,metadata_root)):raise ValueError('Use a separate audit-output directory outside source metadata, extraction and dataset folders')
 ep=locate_export(root,downloads,a.export);export=Source(ep);em=json.loads(export.suffix('manifest.json'));source_meta=rows_sheet(export.suffix('Station_metadata.xlsx'),'Station_metadata');series=rows_sheet(export.suffix('Station_metadata.xlsx'),'Series_metadata')
 if len(source_meta)!=354 or len(series)!=850:raise ValueError('Unexpected selected export version')
 expected_manifest={r['relative_path']:r for r in em['files']};metas={(net(r['network']),provider(net(r['network']),r['source_station_code'])):r for r in source_meta}
 release=locate_package(list(downloads.glob('IHGAR_MeteoSwiss_Station_Update*.zip')),lambda s:json.loads(s.suffix('replay/input_hashes.json')).get('meteoswiss')==EXPORT_SHA and json.loads(s.suffix('replay/input_hashes.json')).get('ihgar')==ORIGINAL_SHA and len(rows_csv(gzip.decompress(s.suffix('Applied_Daily_Changes.csv.gz'))))==183142,a.ihgar_release)
 original_candidates=[root/'working/ihgar_station_update/offline_v3_20261109_205739_593/inputs/IH_GAR_Data.zip']+list(downloads.glob('IH_GAR_Data*.zip'))+[root.parent/'IH_GAR_Data_1/05_Serie_finali_GG']
 original=locate_package(original_candidates,lambda s:(not s.path.is_file() or filehash(s.path)==ORIGINAL_SHA) and len([n for n in s.names() if n.rsplit('/',1)[-1].startswith('PCPD_')])==217,a.ihgar_original)
 if original.path.is_file() and filehash(original.path)!=ORIGINAL_SHA:raise ValueError('Original IH-GAR ZIP differs from the audited archive')
 origmeta={r['nome_file']:r for r in rows_csv(original.suffix('05_Anagrafica_CH_GG.csv'))};original_files={n.rsplit('/',1)[-1]:n for n in original.names() if n.rsplit('/',1)[-1].startswith('PCPD_')}
 release_meta=rows_csv(release.suffix('05_Anagrafica_CH_GG.csv'));links={r['IH-GAR ID']:r for r in release_meta};confirmed={r['stn']:r['IH-GAR ID'] for r in release_meta if r['Output group']=='Updated'};held={r['stn']:r['IH-GAR ID'] for r in release_meta if r['Output group']!='Updated'}
 if len(origmeta)!=217 or len(confirmed)!=124 or len(held)!=2:raise ValueError('IH-GAR crosswalk differs from the reviewed release')
 ledger=rows_csv(gzip.decompress(release.suffix('Applied_Daily_Changes.csv.gz')));actions=defaultdict(Counter)
 for r in ledger:actions[r['IH-GAR ID']][r['Action']]+=1
 reviews=[]
 def issue(severity,identity,subject,reason):reviews.append(dict(severity=severity,station_identity=identity,subject=str(subject),reason=reason,resolution='Open'))
 potential=defaultdict(list)
 selected_meteo={k[1] for k in metas if k[0]=='MeteoSwiss'}
 for ident,r in origmeta.items():
  code=r.get('stn','')
  if code in selected_meteo and code not in confirmed and code not in held:
   potential[code].append(ident);issue('REVIEW','MeteoSwiss:'+code,ident,'Provider code also occurs in IH-GAR outside the qualified update crosswalk; do not count automatically as independent')
 # Registry identities include archive entries and source providers not represented by the crosswalk.
 stations={};family_for={};base_meta={}
 for ident,r in origmeta.items():
  fam='IHGAR:'+ident;link=links.get(ident);code=link['stn'] if link else r.get('stn','');m=metas.get(('MeteoSwiss',code),{})
  status='UPDATED_WITH_ADDITIONS' if actions[ident]['Addition'] else 'CORRECTIONS_ONLY' if actions[ident]['Correction'] else 'ORIGINAL_UNCHANGED'
  stations[fam]=dict(station_identity=fam,network='IHGAR',provider_id=ident,ihgar_original_name=ident,provider_station_name=m.get('station_name',''),matched_meteoswiss_code=code,project_id=(str(m['network_number'])+'_'+str(m['station_number']).zfill(4)) if link and m else '',station_name=link.get('MeteoSwiss name') if link else ident,longitude=r['Longitudine'],latitude=r['Latitudine'],elevation_m=r['Elevazione'],coordinate_basis='ORIGINAL_IHGAR_RETAINED',provider_longitude=m.get('source_longitude',''),provider_latitude=m.get('source_latitude',''),provider_elevation_m=m.get('source_elevation_m',''),coordinate_system='Not explicitly verified in archive metadata',exposure='',identity_status='QUALIFIED_UPDATE_LINK' if code in confirmed else 'HELD_ALIAS_LINK' if code in held else 'ARCHIVE_IDENTITY',update_status=status,added_days=actions[ident]['Addition'],corrected_days=actions[ident]['Correction'],metadata_source=str(original.path)+'!'+original.find('05_Anagrafica_CH_GG.csv'))
  base_meta[('IHGAR',ident)]=r;family_for[('IHGAR',ident)]=fam
 for (n,code),r in metas.items():
  if n=='MeteoSwiss' and code in confirmed:family_for[(n,code)]='IHGAR:'+confirmed[code];continue
  if n=='MeteoSwiss' and code in held:
   family_for[(n,code)]='IHGAR:'+held[code];issue('REVIEW',family_for[(n,code)],code,'Held alias/source relationship: provider product retained as an alternative, not a second independent station');continue
  fam=n+':'+code;family_for[(n,code)]=fam
  stations[fam]=dict(station_identity=fam,network=n,provider_id=code,ihgar_original_name='',provider_station_name=r['station_name'],matched_meteoswiss_code='',project_id=str(r['network_number'])+'_'+str(r['station_number']).zfill(4),station_name=r['station_name'],longitude=r.get('source_longitude'),latitude=r.get('source_latitude'),elevation_m=r.get('source_elevation_m'),coordinate_basis='EXISTING_PROVIDER_METADATA',provider_longitude=r.get('source_longitude'),provider_latitude=r.get('source_latitude'),provider_elevation_m=r.get('source_elevation_m'),coordinate_system=r.get('source_crs') or r.get('source_coordinate_crs'),exposure=r.get('source_station_variant') or '',identity_status='POTENTIAL_ARCHIVE_OVERLAP' if n=='MeteoSwiss' and code in potential else 'SOURCE_PROVIDER_IDENTITY',update_status='NOT_AN_ARCHIVE_UPDATE',added_days=0,corrected_days=0,metadata_source=str(ep)+'!Station_metadata.xlsx')
 # Check the first-stage metadata, without requiring recently deleted stage-two Excel summaries.
 for n,folder in FOLDERS.items():
  p=metadata_root/folder/f'{n}_Metadata.xlsx'
  if not p.is_file():issue('REVIEW',n,p,'First-stage workbook missing; input catalogue metadata are used with explicit provenance');continue
  try:reference={str(r['provider_station_id']):r for r in read_local_rows(p,'Stations')}
  except Exception as e:issue('BLOCKING',n,p,'Cannot read metadata: '+str(e));continue
  for (nn,code),fam in family_for.items():
   if nn!=n:continue
   if n=='MeteoSwiss' and (code in confirmed or code in held):
    m=metas[(n,code)];compare=dict(longitude=m.get('source_longitude'),latitude=m.get('source_latitude'),elevation_m=m.get('source_elevation_m'),project_id=str(m['network_number'])+'_'+str(m['station_number']).zfill(4))
   else:compare=stations[fam]
   if code not in reference:issue('BLOCKING',fam,p,'Identity absent from first-stage metadata: '+code);continue
   ref=reference[code]
   if n=='IHGAR':
    stations[fam]['ihgar_original_name']=ref.get('original_station_name') or code
    if stations[fam]['update_status']=='ORIGINAL_UNCHANGED':stations[fam]['station_name']=stations[fam]['ihgar_original_name']
   for field in ('longitude','latitude','elevation_m'):
    if not identical_metadata(compare.get(field),ref.get(field)):issue('BLOCKING',fam,p,'Metadata conflict in '+field+'; no coordinate substitution made')
   if n!='IHGAR' and str(compare.get('project_id') or '')!=str(ref.get('project_id') or ''):issue('BLOCKING',fam,p,'Assigned project ID differs from first-stage metadata')
 # Expected source bytes are compared against every actual stage-two station product.
 audit_rows=[];manifest=[];file_records=[];counts=Counter()
 def check_file(n,code,folder,name,src,member,daily,interval,basis,role_hint=''):
  fam=family_for[(n,code)];path=stage/FOLDERS[n]/folder/name;baseline=src.read(member);expected_sha=H(baseline)
  stats={};actual_sha='';verified=False
  try:
   actual=path.read_bytes();actual_sha=H(actual)
   if actual_sha!=expected_sha:raise ValueError('Stage-two file differs from source baseline')
   _,stats=grid(actual,daily);verified=True
   first=stats['first_valid'];last=stats['last_valid']
   if first and last:
    if daily:span=(date.fromisoformat(last)-date.fromisoformat(first)).days+1
    else:
     a,b=date.fromisoformat(first),date.fromisoformat(last);span=(b.year-a.year)*12+b.month-a.month+1
    stats['internal_missing_within_observed_span']=span-stats['valid_observations']
   else:stats['internal_missing_within_observed_span']=''
  except Exception as e:issue('BLOCKING',fam,path,str(e))
  record=dict(network=n,provider_id=code,station_identity=fam,time_resolution='daily' if daily else 'monthly',product_folder=folder,filename=name,source_path=str(path),expected_source=str(src.path)+'!'+member,expected_sha256=expected_sha,actual_sha256=actual_sha,copy_status='VERIFIED_BYTE_IDENTICAL' if verified else 'ERROR',accumulation_interval=interval,value_basis=basis,**stats)
  audit_rows.append(record);file_records.append(record)
 for r in series:
  n=net(r['network']);code=provider(n,r['source_station_code']);member=r['output_folder']+'/'+r['output_filename'];baseline=export.read(member)
  if member not in expected_manifest or H(baseline)!=expected_manifest[member]['sha256']:raise ValueError('Export manifest disagreement: '+member)
  nn,folder=DEST[r['output_folder']]
  check_file(n,code,folder,r['output_filename'],export,member,r['temporal_resolution']=='daily',r.get('source_observation_window') or 'Unverified provider interval',r.get('source_value_basis') or '')
  if len(audit_rows)%100==0:print(f'Audited {len(audit_rows)}/1067 source products',flush=True)
 for ident,member in original_files.items():
  link=links.get(ident);updated=link and link['Output group']=='Updated'
  if updated:src=release;name=link['nome_file'];member=release.find('station_files/Updated/'+name)
  else:src=original;name=ident
  check_file('IHGAR',ident,'Daily_updated',name,src,member,True,'Archive intervals are segment-specific or unresolved','Existing updated archive; original coordinates retained')
 if len(audit_rows)!=1067:issue('BLOCKING','All','Source inventory','Expected 1067 selected station-product records, found '+str(len(audit_rows)))
 by_family=defaultdict(list)
 for r in file_records:by_family[r['station_identity']].append(r)
 for r in file_records:
  n=r['network'];code=r['provider_id'];fam=r['station_identity'];daily=r['time_resolution']=='daily';valid=r.get('valid_observations',0);filename=r['filename'];reason='';window06='NOT_APPLICABLE'
  if r['copy_status']!='VERIFIED_BYTE_IDENTICAL':decision='HOLD_FILE_ERROR';reason='Source/copy verification failed'
  elif not valid:decision='EXCLUDE_NO_USABLE_OBSERVATIONS';reason='File has no numeric observations; metadata are retained'
  elif not daily:
   decision='PROPOSE_PARK_MONTHLY';reason='Existing provider monthly product; not newly derived from the selected daily file'
   if n=='MeteoSwiss' and code in held:decision='PARK_MONTHLY_IDENTITY_REVIEW';reason='Held archive/provider alias and different source identity labels require review'
   if n=='MeteoSwiss' and code in potential:decision='HOLD_POTENTIAL_DUPLICATE';reason='Provider code also occurs in the archive outside the qualified crosswalk'
  elif n=='IHGAR':decision='PROPOSE_PRIMARY_DAILY';reason='Archive-first collection; existing changes retained';window06='PENDING_ARCHIVE_SEGMENT_HISTORY'
  elif n=='MeteoSwiss' and (code in confirmed or code in held):decision='KEEP_STAGE2_ALTERNATIVE';reason='Archive record is the primary representation; provider file is not an additional independent gauge';window06='NOT_SELECTED_PRIMARY'
  elif n=='MeteoSwiss' and code in potential:decision='HOLD_POTENTIAL_DUPLICATE';reason='Exact provider-code overlap needs identity review';window06='PENDING_IDENTITY_REVIEW'
  elif n in ('MeteoSwiss','WSL_LWF') and r['product_folder']=='Daily_06UTC':
   decision='PROPOSE_PRIMARY_DAILY';reason='Existing 06 UTC product selected for the proposed collection';window06='PRODUCT_INTERVAL_LABEL_06UTC' if n=='MeteoSwiss' else 'CONDITIONAL_HOURLY_TIMESTAMP_ASSUMPTION'
  elif n in ('MeteoSwiss','WSL_LWF'):decision='KEEP_STAGE2_ALTERNATIVE';reason='Alternative accumulation window retained in stage two';window06='NOT_SELECTED_PRIMARY'
  else:decision='PROPOSE_PRIMARY_DAILY';reason='Provider-native record; not asserted equivalent to 06 UTC';window06='PENDING_PROVIDER_INTERVAL_DOCUMENTATION'
  if daily and decision=='PROPOSE_PRIMARY_DAILY' and any(not populated(stations[fam].get(k)) for k in ('longitude','latitude','elevation_m')):
   decision='PROPOSE_PRIMARY_PENDING_SPATIAL_METADATA';window06='NOT_SPATIALLY_ELIGIBLE';issue('REVIEW',fam,filename,'Valid observations retained, but spatial metadata are incomplete')
  proposed=decision.startswith('PROPOSE_') or decision.startswith('PARK_MONTHLY')
  target=destination/filename if proposed else None
  target_status='NOT_PROPOSED'
  if target:
   if not target.exists():target_status='NOT_YET_COPIED'
   elif target.is_file() and filehash(target)==r['expected_sha256']:target_status='ALREADY_IDENTICAL'
   else:target_status='DESTINATION_CONFLICT';issue('BLOCKING',fam,target,'Existing flat-dataset destination differs; no overwrite proposed')
  manifest.append(dict(station_identity=fam,network=n,provider_id=code,filename=filename,file_naming_rule='Assigned project filename' if NAME.fullmatch(filename) else 'Legacy archive filename retained; no new numbering',time_resolution=r['time_resolution'],decision=decision,reason=reason,common_06utc_status=window06,accumulation_interval=r['accumulation_interval'],source_path=r['source_path'],source_sha256=r['expected_sha256'],valid_observations=valid,first_valid=r.get('first_valid',''),last_valid=r.get('last_valid',''),proposed_destination=str(target) if target else '',destination_status=target_status,monthly_relation='Provider/source monthly product; do not assume sum of primary archive/06 UTC days' if not daily else ''))
 # Build one combined identity register. Counts are record identities, not confirmed physical-site totals.
 for fam,s in stations.items():
  primary=[m for m in manifest if m['station_identity']==fam and m['decision'].startswith('PROPOSE_PRIMARY')]
  parked=[m for m in manifest if m['station_identity']==fam and m['time_resolution']=='monthly' and (m['decision'].startswith('PROPOSE_PARK') or m['decision'].startswith('PARK_MONTHLY'))]
  if len(primary)>1:issue('BLOCKING',fam,'Primary selection','Multiple primary daily files proposed for one identity')
  s['primary_daily_filename']=primary[0]['filename'] if primary else '';s['daily_source']=primary[0]['source_path'] if primary else '';s['daily_interval']=primary[0]['accumulation_interval'] if primary else '';s['common_06utc_status']=primary[0]['common_06utc_status'] if primary else 'NO_PRIMARY_DAILY'
  s['daily_valid_observations']=primary[0]['valid_observations'] if primary else 0;s['daily_first_valid']=primary[0]['first_valid'] if primary else '';s['daily_last_valid']=primary[0]['last_valid'] if primary else ''
  s['parked_monthly_filename']=' | '.join(m['filename'] for m in parked);s['collection_status']=primary[0]['decision'] if primary else 'HELD_OR_METADATA_ONLY'
  if not primary:issue('INFORMATION',fam,'Collection status','No primary daily file proposed; identity retained for review and provenance')
  if s['network']=='IHGAR':
   o=origmeta[s['provider_id']]
   if not all(identical_metadata(s[k],o[ok]) for k,ok in [('longitude','Longitudine'),('latitude','Latitudine'),('elevation_m','Elevazione')]):issue('BLOCKING',fam,'Coordinate preservation','IH-GAR original coordinates/elevation were not retained')
 # Exact equal files across different identities are possible duplicated source series, not automatic deletions.
 content_groups=defaultdict(list)
 for m in manifest:
  if m['decision'].startswith('PROPOSE_PRIMARY'):content_groups[m['source_sha256']].append(m)
 for digest,group in content_groups.items():
  if len({m['station_identity'] for m in group})>1:
   for m in group:issue('REVIEW',m['station_identity'],m['filename'],'Proposed primary daily content is byte-identical to another identity: '+' | '.join(x['station_identity'] for x in group if x is not m)+'; investigate before counting independent gauges')
 dest_groups=defaultdict(list)
 for m in manifest:
  if m['proposed_destination']:dest_groups[m['proposed_destination'].casefold()].append(m)
 for target,group in dest_groups.items():
  if len(group)>1:issue('BLOCKING','All',target,'Multiple source products would occupy one destination filename')
 if destination.is_dir():
  for p in destination.iterdir():
   if p.is_dir():issue('REVIEW','All',p,'Pre-existing nested folder is outside proposed flat dataset; left untouched')
 # Read existing source quality history, not an excuse to alter daily values.
 history=rows_csv(release.suffix('replay/coverage_plan.csv'))
 for r in ledger:
  if r['Action']=='Correction':history.append(dict(id=r['IH-GAR ID'],code='',product=r['Product'],part='Existing value correction',decision='Applied',reason=r['Reason']+'; '+r['Original token']+' -> '+r['Added token'],start=r['Date'],end=r['Date'],days=1))
 issue('SCOPE','All','Audit only','No station data moved, copied, renamed, filled or harmonized. Proposed native collection is not automatically a common-window analysis dataset.')
 issue('SCOPE','All','Common-window readiness','Collection membership is not common-06UTC eligibility. Archive intervals, native provider days, WSL timestamp assumptions and missing spatial metadata remain explicit in Stations/Proposed_files.')
 issue('SCOPE','All','Monthly products','Provider monthly products remain distinct from daily-derived sums. SUPSI exclusions are not applied to other networks.')
 overview=[]
 for n in FOLDERS:
  ss=[s for s in stations.values() if s['network']==n];mm=[m for m in manifest if m['network']==n]
  overview.append(dict(network=n,registry_identities=len(ss),source_product_files=sum(r['network']==n for r in audit_rows),verified_source_files=sum(r['network']==n and r['copy_status']=='VERIFIED_BYTE_IDENTICAL' for r in audit_rows),proposed_primary_daily=sum(m['decision'].startswith('PROPOSE_PRIMARY') for m in mm),proposed_parked_monthly=sum(m['time_resolution']=='monthly' and bool(m['proposed_destination']) for m in mm),alternative_daily_kept_in_stage2=sum(m['time_resolution']=='daily' and m['decision']=='KEEP_STAGE2_ALTERNATIVE' for m in mm),held_files=sum(m['decision'].startswith('HOLD') for m in mm),no_usable_observation_files=sum(m['decision']=='EXCLUDE_NO_USABLE_OBSERVATIONS' for m in mm)))
 audit.mkdir(parents=True,exist_ok=True);book=audit/'Dataset_Register_Audit.xlsx';summary_file=audit/'Dataset_Assembly_Summary.json'
 sheets=[('Overview',overview,list(overview[0])),('Stations',list(stations.values()),['station_identity','network','provider_id','ihgar_original_name','provider_station_name','matched_meteoswiss_code','project_id','station_name','longitude','latitude','elevation_m','coordinate_basis','provider_longitude','provider_latitude','provider_elevation_m','coordinate_system','exposure','identity_status','update_status','added_days','corrected_days','primary_daily_filename','daily_source','daily_interval','common_06utc_status','daily_valid_observations','daily_first_valid','daily_last_valid','parked_monthly_filename','collection_status','metadata_source']),('Proposed_files',manifest,list(manifest[0])),('Source_audit',audit_rows,['network','provider_id','station_identity','time_resolution','product_folder','filename','copy_status','source_path','expected_source','expected_sha256','actual_sha256','accumulation_interval','value_basis','rows','valid_observations','zero_observations','missing_calendar_cells','calendar_padding_cells','internal_missing_within_observed_span','first_valid','last_valid']),('History',history,['id','code','product','part','decision','reason','start','end','days']),('Review',reviews,['severity','station_identity','subject','reason','resolution'])]
 temporary=book.with_name('.Dataset_Register_Audit_'+stamp+'.xlsx')
 try:
  write_book(temporary,sheets)
  if book.exists():
   from zipfile import ZIP_DEFLATED
   backup=audit/f'Previous_Audit_{stamp}.zip'
   with ZipFile(backup,'w',ZIP_DEFLATED) as z:
    z.write(book,book.name)
    if summary_file.exists():z.write(summary_file,summary_file.name)
  os.replace(temporary,book)
 finally:
  if temporary.exists():temporary.unlink()
 blocking=sum(r['severity']=='BLOCKING' for r in reviews);summary=dict(status='AUDIT_COMPLETE_WITH_BLOCKING_ISSUES' if blocking else 'AUDIT_COMPLETE_REVIEW_REQUIRED_BEFORE_TRANSFER',audit_only=True,station_files_moved_or_changed=False,source_products_expected=1067,source_products_audited=len(audit_rows),source_products_verified=sum(r['copy_status']=='VERIFIED_BYTE_IDENTICAL' for r in audit_rows),blocking_issues=blocking,review_items=sum(r['severity']=='REVIEW' for r in reviews),registry_record_identities=len(stations),registry_count_is_not_confirmed_physical_site_count=True,proposed_primary_daily=sum(m['decision'].startswith('PROPOSE_PRIMARY') for m in manifest),proposed_monthly_parking=sum(m['time_resolution']=='monthly' and bool(m['proposed_destination']) for m in manifest),proposal_does_not_mean_common_06utc_eligibility=True,primary_interval_status_counts=dict(Counter(m['common_06utc_status'] for m in manifest if m['decision'].startswith('PROPOSE_PRIMARY'))),ihgar_original_coordinates_preserved=True,ihgar_changed_entries=124,ihgar_entries_with_additions=123,ihgar_corrections_only_entries=1,potential_unqualified_provider_code_links={k:v for k,v in potential.items()},proposed_flat_dataset=str(destination),dataset_created_by_this_run=False,existing_dataset_folder=destination.is_dir(),next_stage_destination_counts=dict(Counter(m['destination_status'] for m in manifest if m['proposed_destination'])),workbook=str(book),input_ihgar_original=str(original.path),original_archive_basis='Audited ZIP SHA-256 verified' if original.path.is_file() else 'Original directory snapshot; file comparisons and first-stage metadata checked, published ZIP hash not available',input_update_release=str(release.path),input_export=str(ep),input_export_sha256=EXPORT_SHA,networks=overview)
 summary_temp=summary_file.with_name('.Dataset_Assembly_Summary_'+stamp+'.json');summary_temp.write_text(json.dumps(summary,indent=2),encoding='utf-8');os.replace(summary_temp,summary_file);print(json.dumps(summary,indent=2));export.close();release.close();original.close()
 return 2 if blocking else 0

if __name__=='__main__':
 try:sys.exit(main())
 except KeyboardInterrupt:raise SystemExit('Audit stopped; source station products were not modified.')

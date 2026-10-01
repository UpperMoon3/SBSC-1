from pathlib import Path
import csv,json,sys,tempfile,shutil
from sbsc_pipeline import run,qc_duplicates
from evaluate import score,read_csv,align
root=Path(__file__).resolve().parent.parent;corpus=root/'corpus'
r=run(str(corpus),None);gt=read_csv(str(corpus/'ground_truth/normalized_journal.csv'))
p2g,_,_=align(r['journal'],gt)
changed=[p2g[i] for i,x in enumerate(r['journal']) if not x['counterparty_name'] and not x['counterparty_identifier'] and x['counterparty_account_or_channel']]
with tempfile.TemporaryDirectory() as t:
 shutil.copytree(corpus/'ground_truth',Path(t)/'ground_truth')
 p=Path(t)/'ground_truth/injection_log.csv';rows=read_csv(str(p));fields=list(rows[0])
 rows=[x for x in rows if not (x['defect_class']=='INCOMPLETE_COUNTERPARTY' and int(x['journal_index']) in changed)]
 with p.open('w') as f:
  w=csv.DictWriter(f,fieldnames=fields);w.writeheader();w.writerows(rows)
 alt=score(r,t)
# Check that cross-file candidates are not called overlaps without overlapping periods.
a=dict(r['journal'][0]); b=dict(a);b['source_file_id']='OTHER';b['statement_period']='2030-01-01/2030-02-01'
d=[];qc_duplicates([a,b],d,True);assert d[0]['code']=='DUP_CANDIDATE_CROSS_FILE_REVIEW'
b['statement_period']=a['statement_period'];d=[];qc_duplicates([a,b],d,True);assert d[0]['code']=='DUP_CANDIDATE_OVERLAP'
b['source_file_id']=a['source_file_id'];d=[];qc_duplicates([a,b],d,True);assert d[0]['code']=='DUP_CANDIDATE_REPEAT'
out={'observed_channel_only_records':len(changed),'alternative_annotation_micro':alt['micro'],'alternative_annotation_counterparty':alt['detection']['INCOMPLETE_COUNTERPARTY'],'duplicate_period_checks':'passed'}
(root/'verified/revision_checks.json').write_text(json.dumps(out,indent=2));print(json.dumps(out,indent=2))

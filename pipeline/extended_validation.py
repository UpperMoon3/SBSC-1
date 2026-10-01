"""Measured scaling and a mapped ETL comparator. Uses synthetic data only."""
import argparse, json, os, platform, resource, statistics, subprocess, sys, time
from pathlib import Path
import openpyxl
from experiments import build_scaled
from sbsc_pipeline import run, read_registry, normalize_file, fill_currency, FIELDS, LINEAGE
from evaluate import read_csv, align, field_equal

HERE=Path(__file__).resolve().parent

def worker(corpus):
    start=time.perf_counter()
    result=run(corpus,None)
    elapsed=time.perf_counter()-start
    peak=resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return {'rows':len(result['journal']),'files':len(result['inventory']),
            'seconds':elapsed,'peak_rss_mb':peak/(2**20 if sys.platform=='darwin' else 1024)}

def mapped_baseline(corpus):
    # Deliberately shares the documented parser to isolate QC's contribution.
    # This is a controlled comparator, not an independent commercial benchmark.
    reg=read_registry(corpus)
    rr={r['source_file_id']:r for r in read_csv(str(Path(corpus)/'statements/request_register.csv'))}
    journal=[]
    for p in sorted((Path(corpus)/'statements').glob('*.xlsx')):
        fid,lay,acct=p.stem.split('_',2)
        r=rr[fid]
        inv={'source_file_id':fid,'client_code':r['client_code'],
             'client_identifier':r['client_identifier'],'declared_account':r['client_account'],
             'declared_period':r['declared_period']}
        records,_=normalize_file(str(p),inv,reg[lay],[])
        journal.extend(records)
    fill_currency(journal)
    full=run(corpus,None)
    assert len(journal)==len(full['journal'])
    same=sum(all(a.get(f)==b.get(f) for f in FIELDS) for a,b in zip(journal,full['journal']))
    return {'rows':len(journal),'rows_identical_to_full_on_source_fields':same,
            'rows_with_source_lineage':sum(all(r.get(f) is not None for f in LINEAGE[:4]) for r in journal),
            'qc_candidates':0,'hash_verification':False,'source_row_retention':True,
            'provenance':'shared parser and mapping rules; omits QC and hash control'}

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--corpus',required=True);p.add_argument('--worker',action='store_true');p.add_argument('--out',default='extended_results');a=p.parse_args()
    if a.worker:
        print(json.dumps(worker(a.corpus)));sys.exit()
    out=Path(a.out).resolve();out.mkdir(parents=True,exist_ok=True)
    (out/'mapped_baseline.json').write_text(json.dumps(mapped_baseline(a.corpus),indent=2))
    cache=out/'scale_inputs';cache.mkdir(exist_ok=True)
    results=[]
    for k in [1,8,42,84]:
        corpus=build_scaled(str(Path(a.corpus).resolve()),k,str(cache))
        trials=[]
        for rep in range(3):
            result=subprocess.run([sys.executable,str(Path(__file__).resolve()),'--worker','--corpus',corpus],capture_output=True,text=True,check=True)
            trials.append(json.loads(result.stdout))
        secs=[r['seconds'] for r in trials]
        row={'rows':trials[0]['rows'],'files':trials[0]['files'],'runs':3,
             'median_s':statistics.median(secs),'min_s':min(secs),'max_s':max(secs),
             'peak_rss_mb':max(r['peak_rss_mb'] for r in trials),'raw_runs':trials}
        results.append(row)
        data={'sizes':results,'machine':{'platform':platform.platform(),'python':platform.python_version(),'openpyxl':openpyxl.__version__},
              'protocol':'Three fresh processes per size; wall time includes run() cleanup; generation, interpreter startup and figure rendering excluded; warm OS caches permitted. Population grows by adding independently seeded synthetic corpora, not longer histories within one account.'}
        (out/'scaling_extended.json').write_text(json.dumps(data,indent=2))
        print(json.dumps(row),flush=True)

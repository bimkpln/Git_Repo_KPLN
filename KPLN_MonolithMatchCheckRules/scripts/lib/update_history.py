"""Prepare a reviewable history transaction; never overwrite the registry early.

After both reports are verified against canonical_entries.json, commit with
--commit. File hashes in the report verification receipts must match the same
canonical hierarchy. The registry is protected from concurrent lost updates.
"""
import argparse,copy,hashlib,json,os,shutil
from pathlib import Path
from history import reconcile,export_history,atomic_write

def filehash(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def read(p):return json.loads(Path(p).read_text(encoding='utf8'))

def prepare(registry_path,entries_path,coverage_path,out,run_id,date,certificates_path=None):
    registry_path=Path(registry_path);out=Path(out)
    reg=read(registry_path);entries=read(entries_path);physical=[g for e in entries for g in e.get('local_instances') or [e]]
    result=reconcile(reg,physical,run_id,date,read(coverage_path),read(certificates_path) if certificates_path else [])
    if not result['complete']:atomic_write(out/'ambiguous_history.json',result);raise ValueError('Resolve ambiguous feature matches before report generation')
    canonical,newgroups=export_history(entries,result['registry'],result['physical_uid_map'])
    for g in result['registry']['groups'].values():g['active']=False
    result['registry']['groups'].update(newgroups)
    atomic_write(out/'canonical_entries.json',canonical)
    atomic_write(out/'registry_pending.json',result['registry'])
    atomic_write(out/'transaction.json',{'registry_path':str(registry_path.resolve()),'previous_sha256':filehash(registry_path),
        'pending_sha256':filehash(out/'registry_pending.json'),'entries_sha256':filehash(out/'canonical_entries.json'),'run_id':run_id})

def commit(out,plugin_receipt,coordinator_receipt):
    out=Path(out);t=read(out/'transaction.json');dest=Path(t['registry_path']);lock=dest.with_suffix('.lock')
    # Exclusive lock file; never replace a concurrently updated registry.
    with lock.open('x') as stream:stream.write(str(os.getpid()))
    try:
        if filehash(dest)!=t['previous_sha256'] or filehash(out/'registry_pending.json')!=t['pending_sha256'] or filehash(out/'canonical_entries.json')!=t['entries_sha256']:raise ValueError('Stale or modified transaction')
        for p in [plugin_receipt,coordinator_receipt]:
            r=read(p)
            if r.get('passed') is not True or r.get('canonical_entries_sha256')!=t['entries_sha256']:raise ValueError('Both reports must verify against the same hierarchy')
            artifacts=r.get('artifact_hashes',{})
            if not artifacts or not all((Path(p).parent/n).is_file() and filehash(Path(p).parent/n)==h for n,h in artifacts.items()):raise ValueError('Stale report verification')
        backup=dest.parent/'backups'/t['previous_sha256'];backup.mkdir(parents=True,exist_ok=True)
        shutil.copy2(dest,backup/'registry.json');atomic_write(dest,read(out/'registry_pending.json'))
    finally:lock.unlink()

if __name__=='__main__':
    p=argparse.ArgumentParser();sub=p.add_subparsers(dest='command',required=True)
    a=sub.add_parser('prepare');a.add_argument('registry');a.add_argument('entries');a.add_argument('coverage');a.add_argument('out');a.add_argument('run_id');a.add_argument('date');a.add_argument('--certificates')
    a=sub.add_parser('commit');a.add_argument('out');a.add_argument('plugin_receipt');a.add_argument('coordinator_receipt');a=p.parse_args()
    if a.command=='prepare':prepare(a.registry,a.entries,a.coverage,a.out,a.run_id,a.date,a.certificates)
    else:commit(a.out,a.plugin_receipt,a.coordinator_receipt)

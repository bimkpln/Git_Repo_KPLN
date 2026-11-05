"""Prepare provenance only; this command does NOT calculate NWD differences."""
import argparse,datetime,hashlib,json,os,shutil,sys
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]

def filehash(path):
    h=hashlib.sha256()
    with Path(path).open('rb') as f:
        for chunk in iter(lambda:f.read(4*1024*1024),b''):h.update(chunk)
    return h.hexdigest()

def inventory(root):
    files=[p for folder in ['rules','scripts','tests','skills','examples','docs','engine'] for p in (root/folder).rglob('*')
           if p.is_file() and '__pycache__' not in p.parts and p.suffix!='.pyc']
    files += [root/name for name in ['VERSION','README.md','AGENTS.md','CHANGELOG.md','.gitignore']]
    return {p.relative_to(root).as_posix():filehash(p) for p in sorted(files)}

def prepare(config_path,nwd_path,output,operator,results_root=None):
    if (ROOT/'local/rules-update.lock').exists():raise ValueError('Shared rules update in progress; retry after it completes')
    config_path=Path(config_path).resolve();nwd_path=Path(nwd_path).resolve()
    if (output is None)==(results_root is None):raise ValueError('Choose exact output or results root')
    if not operator.strip():raise ValueError('Operator required')
    if not nwd_path.is_file() or nwd_path.suffix.lower()!='.nwd':raise ValueError('Existing NWD required')
    cfg=json.loads(config_path.read_text(encoding='utf-8-sig'))
    started=datetime.datetime.now(datetime.timezone.utc);stamp=started.isoformat()
    before=nwd_path.stat();nwd_hash=filehash(nwd_path);after=nwd_path.stat()
    if (before.st_size,before.st_mtime_ns)!=(after.st_size,after.st_mtime_ns):raise ValueError('NWD changed during hashing')
    for view in cfg.get('views',[]):
        p=Path(view['source']);p=p if p.is_absolute() else config_path.parent/p
        if not p.is_file() or filehash(p)!=view['source_sha256']:raise ValueError('View source absent or hash differs; register actual drawings')
        view['source']=str(p.resolve())
    manifest=inventory(ROOT)
    if results_root is not None:
        sys.path.insert(0,str(ROOT/'scripts/lib'))
        from result_directory import allocate
        output=allocate(results_root,cfg,started)
        # Subsequent imports must use the pinned snapshot, including this helper.
        sys.modules.pop('result_directory',None);sys.path.pop(0)
    else:
        output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    try:
        snapshot=output/'rules_snapshot'
        for name,value in manifest.items():
            target=snapshot/name;target.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(ROOT/name,target)
            if filehash(target)!=value:raise ValueError('Rules changed while snapshotting')
        if (ROOT/'local/rules-update.lock').exists() or inventory(ROOT)!=manifest:raise ValueError('Rules changed during preparation; retry in a new run')
        # Load policy from this immutable snapshot, not a later shared edit.
        sys.path.insert(0,str(snapshot/'scripts/lib'))
        from policy import resolve,digest
        from registry_scope import registry_path
        state=Path(cfg.get('state_root',ROOT/'local/shared-state'))
        if not state.is_absolute():state=config_path.parent/state
        cfg.update(input_nwd_sha256=nwd_hash,nwd_path=str(nwd_path),state_root=str(state.resolve()))
        profile=resolve(cfg)
        run={**cfg,'run_id':output.name,'operator':operator,'prepared_at_utc':stamp,
            'result_directory':str(output),'result_date_basis':'run_start_Europe/Moscow',
            'status':'prepared_not_checked','geometry_engine_status':'integration_required',
            'rules_version':(snapshot/'VERSION').read_text().strip(),'rules_manifest_sha256':digest(manifest),
            'profile':profile,'registry_path':str(registry_path(state/'registries',cfg)),
            'config_source_sha256':filehash(config_path)}
        (output/'rules_manifest.json').write_text(json.dumps(manifest,indent=2),encoding='utf8')
        (output/'run.json').write_text(json.dumps(run,ensure_ascii=False,indent=2),encoding='utf8')
        print(json.dumps({'status':run['status'],'run':str(output/'run.json'),'registry':run['registry_path']},ensure_ascii=False))
    except Exception as error:
        (output/'PREPARATION_FAILED.txt').write_text(str(error),encoding='utf8')
        raise

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__)
    for arg in ['config','nwd','operator']:p.add_argument('--'+arg,required=True)
    target=p.add_mutually_exclusive_group(required=True)
    target.add_argument('--output');target.add_argument('--results-root')
    a=p.parse_args();prepare(a.config,a.nwd,a.output,a.operator,a.results_root)

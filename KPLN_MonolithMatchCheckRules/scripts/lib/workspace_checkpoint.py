"""Immutable, resumable workspace checkpoints. Standard library, no model edits.

Build only at a quiescent stage boundary. READY is written after remote hashes
are verified. Work locally; never execute calculations on the network replica.
"""
import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import shutil
import sys
import uuid
import zipfile

def digest(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for block in iter(lambda: f.read(4*1024*1024), b''):
            h.update(block)
    return h.hexdigest()

def write_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + '.' + uuid.uuid4().hex + '.tmp')
    tmp.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')
    os.replace(tmp, path)

def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))

def relative(value):
    p = PurePosixPath(value)
    if not value or p.is_absolute() or '\\' in value or any(x in ('..', '.', '') or ':' in x for x in p.parts):
        raise ValueError('Unsafe relative path: ' + value)
    if value != p.as_posix():
        raise ValueError('Noncanonical path: ' + value)
    return p

def inside(root, name):
    p = root.joinpath(*relative(name).parts)
    if not p.resolve().is_relative_to(root.resolve()):
        raise ValueError('Path escapes root: ' + name)
    return p

def select(root, cfg):
    excluded = cfg.get('exclude_prefixes', [])
    result = {}
    for name in cfg['include']:
        base = inside(root, name)
        if not base.exists():
            raise FileNotFoundError(base)
        paths = [base] if base.is_file() else base.rglob('*')
        for p in paths:
            rel = p.relative_to(root).as_posix()
            if any(rel == x or rel.startswith(x.rstrip('/')+'/') for x in excluded):
                continue
            if any(x in ('.git', '__pycache__', 'node_modules') for x in p.relative_to(root).parts):
                continue
            if p.suffix in ('.pyc', '.tmp', '.lock') or not p.is_file():
                continue
            if p.is_symlink() or not p.resolve().is_relative_to(root):
                raise ValueError('External link: ' + rel)
            result[rel] = p
    return dict(sorted(result.items()))

def build(root, destination, config, baseline=None, part_mb=256):
    root = Path(root).resolve()
    dest = Path(destination).resolve()
    if dest.exists():
        raise FileExistsError('Use a new checkpoint directory: ' + str(dest))
    cfg = read(config)
    selected = select(root, cfg)
    old = read(Path(baseline)/'manifest.json') if baseline else {}
    if old and (old['workspace_id'] != cfg['workspace_id'] or old['selection'] != cfg):
        raise ValueError('Baseline belongs to another workspace/selection')
    if old and not (Path(baseline)/'published.json').exists():
        raise ValueError('Baseline must already be published')
    previous = old.get('files', {})
    dest.mkdir(parents=True)
    (dest/'objects').mkdir()
    rows, new_rows, reused = {}, [], 0
    # Hash every file; do not assume unchanged mtime implies unchanged content.
    for i, (name, p) in enumerate(selected.items(), 1):
        before = p.stat()
        h = digest(p)
        after = p.stat()
        if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
            raise RuntimeError('Source changed while hashing: ' + name)
        row = {'sha256': h, 'size': after.st_size, 'mtime_ns': after.st_mtime_ns}
        if name in previous and previous[name]['sha256'] == h:
            row.update({k:previous[name][k] for k in ('archive', 'member')})
            reused += 1
        else:
            new_rows.append((name, p, row))
        rows[name] = row
        if i % 20000 == 0:
            print('HASH', i, '/', len(selected), flush=True)
    # Identical source bytes are stored once; restore creates independent files.
    packs, items, size = [], [], 0
    content = {}
    for name, p, row in new_rows:
        key = (row['sha256'], row['size'])
        if key in content:
            continue
        content[key] = name
        if items and size + row['size'] > part_mb*1024*1024:
            packs.append(items); items = []; size = 0
        items.append((name, p, row)); size += row['size']
    if items:
        packs.append(items)
    objects = dict(old.get('objects', {}))
    location = {}
    for n, batch in enumerate(packs, 1):
        tmp = dest/'objects'/('part-%06d.zip.tmp' % n)
        with zipfile.ZipFile(tmp, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=1, allowZip64=True) as z:
            for name, p, row in batch:
                h = hashlib.sha256()
                info = zipfile.ZipInfo(name, date_time=(1980,1,1,0,0,0))
                info.compress_type = zipfile.ZIP_DEFLATED
                info._compresslevel = 1
                with p.open('rb') as src, z.open(info, 'w', force_zip64=True) as out:
                    for b in iter(lambda: src.read(4*1024*1024), b''):
                        h.update(b); out.write(b)
                if h.hexdigest() != row['sha256']:
                    raise RuntimeError('Source changed while packing: ' + name)
        archive = digest(tmp)
        target = dest/'objects'/(archive+'.zip')
        os.replace(tmp, target)
        objects[archive] = {'bytes':target.stat().st_size}
        for name,p,row in batch:
            location[(row['sha256'],row['size'])] = (archive,name)
        print('PACK', n, '/', len(packs), target.stat().st_size, flush=True)
    for name,p,row in new_rows:
        row['archive'],row['member'] = location[(row['sha256'],row['size'])]
    # Detect changes during the whole checkpoint, including files initially reused.
    if set(select(root,cfg)) != set(selected):
        raise RuntimeError('Source file list changed during checkpoint')
    for name,p in selected.items():
        st = p.stat()
        if (st.st_size,st.st_mtime_ns) != (rows[name]['size'],rows[name]['mtime_ns']):
            raise RuntimeError('Source changed during checkpoint: ' + name)
    needed = {v['archive'] for v in rows.values()}
    manifest = {'schema':1, 'checkpoint_id':dest.name, 'workspace_id':cfg['workspace_id'],
                'created_at_utc':dt.datetime.now(dt.timezone.utc).isoformat(),
                'source_root':str(root), 'selection':cfg, 'files':rows,
                'objects':{k:objects[k] for k in sorted(needed)},
                'baseline_checkpoint':old.get('checkpoint_id'),
                'file_count':len(rows),'logical_bytes':sum(v['size'] for v in rows.values()),
                'reused_files':reused, 'complete_reports':False,
                'registry_copies_are_archival':True}
    write_json(dest/'manifest.json',manifest)
    write_json(dest/'BUILT.json',{'manifest_sha256':digest(dest/'manifest.json')})
    print('BUILT',dest,len(rows),'files; reused',reused,flush=True)
    return manifest

def verified_copy(src, dest, expected):
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        if digest(dest) != expected:
            raise RuntimeError('Existing object is corrupt; never overwrite: '+str(dest))
        return
    tmp = dest.with_name(dest.name+'.'+uuid.uuid4().hex+'.partial')
    shutil.copyfile(src,tmp)
    if digest(tmp) != expected:
        raise RuntimeError('Copy verification failed: '+str(tmp))
    # Unique immutable paths; rename on Windows refuses an existing destination.
    try:
        tmp.rename(dest)
    except FileExistsError:
        if digest(dest) != expected:
            raise
        tmp.unlink()

def publish(checkpoint, remote):
    cp, remote = Path(checkpoint).resolve(), Path(remote)
    m = read(cp/'manifest.json')
    h = digest(cp/'manifest.json')
    if read(cp/'BUILT.json')['manifest_sha256'] != h:
        raise RuntimeError('Local manifest changed')
    dest = remote/'checkpoints'/m['checkpoint_id']
    relative(m['checkpoint_id'])
    if len(PurePosixPath(m['checkpoint_id']).parts) != 1:
        raise ValueError('Invalid checkpoint ID')
    for i,key in enumerate(m['objects'],1):
        if len(key)!=64 or any(c not in '0123456789abcdef' for c in key):
            raise ValueError('Invalid object digest')
        obj = remote/'objects'/key[:2]/(key+'.zip')
        local = cp/'objects'/(key+'.zip')
        if not local.exists() and not obj.exists():
            raise FileNotFoundError('Missing baseline object: '+key)
        verified_copy(local,obj,key)
        print('SYNC',i,'/',len(m['objects']),flush=True)
    verified_copy(cp/'manifest.json',dest/'manifest.json',h)
    marker = {'manifest_sha256':h,'file_count':m['file_count'],
              'verified_at_utc':dt.datetime.now(dt.timezone.utc).isoformat(),
              'complete_reports':False}
    if (dest/'READY.json').exists():
        if read(dest/'READY.json')['manifest_sha256'] != h:
            raise RuntimeError('Checkpoint name collision')
    else:
        write_json(dest/'READY.json',marker)
    write_json(cp/'published.json',{'remote':str(remote),'checkpoint':str(dest),**marker})
    print('PUBLISHED',dest,flush=True)
    return dest

def restore(remote, checkpoint_id, target, prefixes=None):
    remote,target = Path(remote),Path(target).resolve()
    relative(checkpoint_id)
    if len(PurePosixPath(checkpoint_id).parts)!=1:
        raise ValueError('Invalid checkpoint ID')
    cp = remote/'checkpoints'/checkpoint_id
    if digest(cp/'manifest.json') != read(cp/'READY.json')['manifest_sha256']:
        raise RuntimeError('Unverified manifest')
    m = read(cp/'manifest.json')
    files = {k:v for k,v in m['files'].items() if not prefixes or any(k==p or k.startswith(p.rstrip('/')+'/') for p in prefixes)}
    if not files:
        raise ValueError('Empty selection')
    for name,row in files.items():
        relative(name); relative(row['member'])
        if row['archive'] not in m['objects']:
            raise ValueError('Unknown archive')
        inside(target,name)
    if target.exists() and any(target.iterdir()):
        raise FileExistsError('Restore only to a new/empty directory')
    target.mkdir(parents=True,exist_ok=True)
    groups={}
    for name,row in files.items():
        groups.setdefault(row['archive'],[]).append((name,row))
    for key,batch in groups.items():
        p=remote/'objects'/key[:2]/(key+'.zip')
        if digest(p)!=key:
            raise RuntimeError('Archive verification failed')
        with zipfile.ZipFile(p) as z:
            for name,row in batch:
                dest=inside(target,name)
                dest.parent.mkdir(parents=True,exist_ok=True)
                with z.open(row['member']) as src,dest.open('xb') as out:
                    shutil.copyfileobj(src,out,1024*1024)
                if dest.stat().st_size!=row['size'] or digest(dest)!=row['sha256']:
                    raise RuntimeError('Restored file mismatch: '+name)
    write_json(target/'RESTORED_CHECKPOINT.json',{'checkpoint':checkpoint_id,'file_count':len(files),'prefixes':prefixes,'source':str(cp),'registry_copies_are_archival':True})
    print('RESTORED',len(files),'files',flush=True)
    return len(files)

def main():
    p=argparse.ArgumentParser(description=__doc__)
    s=p.add_subparsers(dest='command',required=True)
    b=s.add_parser('build');b.add_argument('--root',required=True);b.add_argument('--out',required=True);b.add_argument('--config',required=True);b.add_argument('--baseline');b.add_argument('--part-mb',type=int,default=256)
    u=s.add_parser('publish');u.add_argument('--checkpoint',required=True);u.add_argument('--remote',required=True)
    r=s.add_parser('restore');r.add_argument('--remote',required=True);r.add_argument('--id',required=True);r.add_argument('--out',required=True);r.add_argument('--prefix',action='append')
    a=p.parse_args()
    if a.command=='build':build(a.root,a.out,a.config,a.baseline,a.part_mb)
    elif a.command=='publish':publish(a.checkpoint,a.remote)
    else:restore(a.remote,a.id,a.out,a.prefix)

if __name__=='__main__':main()

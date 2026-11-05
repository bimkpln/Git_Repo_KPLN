"""Write import admission and material preflight; never issue final reports."""
import argparse,json,sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parent/'lib'))
from nwd_import import validate_export
from preflight import prepare

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--config',required=True);p.add_argument('--export',required=True);p.add_argument('--output',required=True)
    a=p.parse_args();config=json.loads(Path(a.config).read_text(encoding='utf-8-sig'))
    fragments,receipt=validate_export(a.export,config['input_nwd_sha256'])
    result=prepare(config,fragments);dest=Path(a.output);dest.mkdir(parents=True,exist_ok=False)
    for name,data in [('fragment_inventory.json',fragments),('import_validation.json',receipt),('selection_preflight.json',result)]:
        (dest/name).write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8')
    print(json.dumps({'import_complete':receipt['import_complete'],'materials':result['material_selection']['counts'],
        'complete_geometry_check':False},ensure_ascii=False))
    if not receipt['import_complete'] or not result['ready_for_geometry']:raise SystemExit(2)

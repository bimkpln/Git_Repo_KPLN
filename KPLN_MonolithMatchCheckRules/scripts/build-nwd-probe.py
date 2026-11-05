"""Compile the checked-in exporter with installed Navisworks .NET assemblies."""
import argparse,hashlib,json,os,subprocess
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()

def build(runtime,output):
    runtime=Path(runtime).resolve();output=Path(output).resolve()
    source=ROOT/'engine/NwdFragmentProbe.cs'
    refs=[runtime/n for n in ['Autodesk.Navisworks.Api.dll','Autodesk.Navisworks.Controls.dll',
        'Autodesk.Navisworks.ComApi.dll','Autodesk.Navisworks.Interop.ComApi.dll']]
    compiler=Path(os.environ.get('WINDIR','C:/Windows'))/'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    if not compiler.is_file() or not all(p.is_file() for p in refs):raise ValueError('Installed compiler/Navisworks assemblies required')
    output.mkdir(parents=True,exist_ok=False)
    exe=output/'NwdFragmentProbe.exe'
    subprocess.run([str(compiler),'/nologo','/platform:x64','/optimize+','/r:System.Web.Extensions.dll',
        *['/r:'+str(p) for p in refs],'/out:'+str(exe),str(source)],check=True)
    exe.with_suffix('.exe.config').write_text('<configuration><startup useLegacyV2RuntimeActivationPolicy="true">'
        '<supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.8" /></startup></configuration>',encoding='utf8')
    receipt={'source_sha256':sha(source),'executable_sha256':sha(exe),'runtime':str(runtime),
        'reference_hashes':{p.name:sha(p) for p in refs},'compiler_sha256':sha(compiler),
        'geometry_check_performed':False}
    (output/'build_receipt.json').write_text(json.dumps(receipt,indent=2),encoding='utf8')
    return receipt

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--runtime',required=True);p.add_argument('--output',required=True);a=p.parse_args()
    print(json.dumps(build(a.runtime,a.output)))

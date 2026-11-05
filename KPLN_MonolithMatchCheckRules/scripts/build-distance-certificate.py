"""Build the standalone read-only surface distance proof; not a full check."""
import argparse,subprocess,hashlib,json
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]

def build(output):
    output=Path(output).resolve()
    if output.exists():raise ValueError('Choose a new executable path; earlier builds are preserved')
    source=ROOT/'engine/SurfaceCertificate.cs';compiler=Path(r'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe')
    if not compiler.is_file():raise FileNotFoundError('Windows .NET Framework C# compiler is required')
    output.parent.mkdir(parents=True,exist_ok=True)
    subprocess.run([str(compiler),'/nologo','/optimize+','/out:'+str(output),str(source)],check=True)
    receipt={'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'executable_sha256':hashlib.sha256(output.read_bytes()).hexdigest(),
        'purpose':'bounded whole-triangle surface distance proof','complete_matching_engine':False}
    output.with_suffix('.build.json').write_text(json.dumps(receipt,indent=2),encoding='utf8')
    return output

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--output',required=True);a=p.parse_args();print(build(a.output))

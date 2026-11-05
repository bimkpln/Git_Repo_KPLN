"""Optional literal binary matrices: JSON matrices are display representations."""
from pathlib import Path
import hashlib,re,struct,math

NAME='nwd_fragment_matrices.bin'
RECORD=struct.Struct('<ii16d')

def filehash(path):
    h=hashlib.sha256()
    with Path(path).open('rb') as f:
        for block in iter(lambda:f.read(4*1024*1024),b''):h.update(block)
    return h.hexdigest()

class LiteralMatrices:
    def __init__(self,directory,manifest):
        self.path=Path(directory)/NAME;self.stream=None;self.records=0;self.changed=0;self.maximum=0.
        name=manifest.get('literal_matrix_file');sha=manifest.get('literal_matrix_sha256')
        if name is None and sha is None and not self.path.exists():return
        if name!=NAME or not isinstance(sha,str) or not re.fullmatch('[0-9a-f]{64}',sha):
            raise ValueError('Fixed literal matrix file name and SHA256 required')
        if not self.path.is_file() or self.path.stat().st_size%RECORD.size or filehash(self.path)!=sha:
            raise ValueError('Missing, truncated or changed literal matrices')
        self.sha256=sha;self.stream=self.path.open('rb')
    def node(self,node):
        if self.stream is None:return node
        for index,f in enumerate(node['fragmentRecords']):
            data=self.stream.read(RECORD.size)
            if len(data)!=RECORD.size:raise ValueError('Missing literal matrix record')
            owner,fragment,*matrix=RECORD.unpack(data)
            if owner!=node['index'] or fragment!=index or f['fragment_id']!=str(owner)+':'+str(fragment):
                raise ValueError('Literal matrix owner/order mismatch')
            if not all(math.isfinite(x) for x in matrix):raise ValueError('Nonfinite literal matrix')
            displayed=f['local_to_world']
            if len(displayed)!=16 or any(isinstance(x,bool) or not isinstance(x,(int,float)) or not math.isfinite(x) for x in displayed):
                raise ValueError('Invalid display matrix')
            if displayed!=matrix:
                self.changed+=1;self.maximum=max(self.maximum,max(abs(a-b) for a,b in zip(displayed,matrix)))
            f['local_to_world']=matrix;self.records+=1
        return node
    def finish(self):
        if self.stream is None:return None
        if self.stream.read(1):raise ValueError('Trailing unaccounted literal matrices')
        return {'passed':True,'file':NAME,'sha256':self.sha256,'records':self.records,
            'JSON_display_matrices_with_differences':self.changed,
            'maximum_JSON_display_matrix_component_difference':self.maximum,
            'authoritative_transform':'binary_float64','JSON_matrix_is_display_only':True}
    def close(self):
        if self.stream is not None:self.stream.close()

"""Iterate a large JSON mapping with memory bounded by one entry and a chunk."""
import json

def entries(path,chunk_size=1<<20):
    decoder=json.JSONDecoder()
    with open(path,encoding='utf-8-sig') as source:
        buffer='';position=0;ended=False
        def more():
            nonlocal buffer,position,ended
            chunk=source.read(chunk_size)
            buffer=buffer[position:]+chunk;position=0
            if not chunk:ended=True
        def token():
            nonlocal position
            while True:
                while position<len(buffer) and buffer[position].isspace():position+=1
                if position<len(buffer):return buffer[position]
                if ended:raise ValueError('Unexpected end of JSON mapping')
                more()
        def value():
            nonlocal position
            token()
            while True:
                try:
                    result,end=decoder.raw_decode(buffer,position)
                    # A number can continue in the next input chunk.
                    if end==len(buffer) and not ended:more();continue
                    if type(result) in [int,float] and end<len(buffer) and buffer[end] not in ',}] \t\r\n':
                        if ended:raise ValueError('Malformed JSON number')
                        more();continue
                    position=end;return result
                except json.JSONDecodeError:
                    if ended:raise
                    more()
        if token()!='{':raise ValueError('JSON root must be a mapping')
        position+=1;first=True
        while True:
            current=token()
            if current=='}':
                position+=1
                if (buffer[position:]+source.read()).strip():raise ValueError('Data follows JSON mapping')
                return
            if not first:
                if current!=',':raise ValueError('JSON mapping separator missing')
                position+=1
            key=value()
            if not isinstance(key,str):raise ValueError('JSON mapping key must be a string')
            if token()!=':':raise ValueError('JSON mapping colon missing')
            position+=1;item=value();yield key,item;first=False

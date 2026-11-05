"""Iterate a large JSON array while retaining only one decoded value."""
import json

def items(path,chunk_size=1<<20):
    decoder=json.JSONDecoder()
    with open(path,encoding='utf-8-sig') as source:
        buffer='';position=0;ended=False
        def more():
            nonlocal buffer,position,ended
            chunk=source.read(chunk_size);buffer=buffer[position:]+chunk;position=0
            if not chunk:ended=True
        def token():
            nonlocal position
            while True:
                while position<len(buffer) and buffer[position].isspace():position+=1
                if position<len(buffer):return buffer[position]
                if ended:raise ValueError('Unexpected end of JSON array')
                more()
        def value():
            nonlocal position
            token()
            while True:
                try:
                    result,end=decoder.raw_decode(buffer,position)
                    if end==len(buffer) and not ended:more();continue
                    if type(result) in [int,float] and end<len(buffer) and buffer[end] not in ',}] \t\r\n':
                        if ended:raise ValueError('Malformed JSON number')
                        more();continue
                    position=end;return result
                except json.JSONDecodeError:
                    if ended:raise
                    more()
        if token()!='[':raise ValueError('JSON root must be an array')
        position+=1;first=True
        while True:
            current=token()
            if current==']':
                position+=1
                if (buffer[position:]+source.read()).strip():raise ValueError('Data follows JSON array')
                return
            if not first:
                if current!=',':raise ValueError('JSON array separator missing')
                position+=1
            yield value();first=False

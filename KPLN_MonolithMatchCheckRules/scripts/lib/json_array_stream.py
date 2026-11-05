"""Strict incremental JSON arrays; bounds memory by one item and input chunk."""
import json

def array_items(path,wrapper_key=None,chunk_size=1<<20):
    if isinstance(chunk_size,bool) or not isinstance(chunk_size,int) or chunk_size<1:
        raise ValueError('Positive chunk size required')
    decoder=json.JSONDecoder()
    with open(path,encoding='utf-8-sig') as stream:
        buf='';position=0;ended=False
        def more():
            nonlocal buf,position,ended
            block=stream.read(chunk_size);buf=buf[position:]+block;position=0
            if not block:ended=True
        def token():
            nonlocal position
            while True:
                while position<len(buf) and buf[position].isspace():position+=1
                if position<len(buf):return buf[position]
                if ended:raise ValueError('Unexpected end of JSON array')
                more()
        def expect(character):
            nonlocal position
            if token()!=character:raise ValueError('Expected JSON delimiter '+character)
            position+=1
        def value():
            nonlocal position
            token()
            while True:
                try:
                    result,end=decoder.raw_decode(buf,position)
                    if end==len(buf) and not ended:more();continue
                    if end<len(buf) and buf[end] not in ',]} \t\r\n:':
                        if ended:raise ValueError('Invalid token after JSON value')
                        more();continue
                    position=end;return result
                except json.JSONDecodeError as error:
                    if ended:raise ValueError('Malformed JSON array') from error
                    more()
        if wrapper_key is not None:
            expect('{')
            if value()!=wrapper_key:raise ValueError('Unexpected JSON array wrapper key')
            expect(':')
        expect('[')
        if token()!=']':
            while True:
                yield value()
                if token()==']':break
                expect(',')
                if token()==']':raise ValueError('Trailing JSON array comma')
        expect(']')
        if wrapper_key is not None:expect('}')
        if (buf[position:]+stream.read()).strip():raise ValueError('Trailing JSON data')

"""Metadata-only Groups invitation mail cleanup; never reads bodies or links.

Only eventMessage records from the exact Groups domain are eligible. Unknown
ordinary senders are not filtered using the contacts policy. Matches move to
Junk Email (system-decided unwanted; links disabled, recoverable).
"""
import argparse
import json
import re
import urllib.parse
from outlook import config
from outlook import outlook_personal_graph as graph

SELECT='id,from'
DESTINATION='junkemail'

def eligible(message):
    address=((message.get('from') or {}).get('emailAddress') or {}).get('address')
    return (message.get('@odata.type') in ('#microsoft.graph.eventMessage', '#microsoft.graph.eventMessageRequest', '#microsoft.graph.eventMessageResponse')
            and isinstance(address,str) and re.fullmatch(r'[^\s@<>]+@groups\.outlook\.com',address,re.I) is not None)

def scan(apply=False):
    expected_account=config.ACCOUNT
    if not expected_account:
        raise RuntimeError('ACCOUNT_NOT_CONFIGURED')
    me=graph.graph_request('GET','/me',query={'$select':'mail,userPrincipalName'})
    if not any(str(me.get(k,'')).casefold()==expected_account for k in ('mail','userPrincipalName')):
        raise RuntimeError('ACCOUNT_MISMATCH')
    path='/me/mailFolders/inbox/messages?$select='+SELECT+'&$top=100'
    ids=[];seen=set();checked=0
    while path:
        if path in seen: raise RuntimeError('PAGINATION_LOOP')
        if path.startswith('https://') and not path.startswith(graph.GRAPH_BASE+'/me/'):
            raise RuntimeError('INVALID_API_URL')
        seen.add(path)
        page=graph.graph_request('GET',path)
        for message in page.get('value',[]):
            checked+=1
            if eligible(message): ids.append(message['id'])
        path=page.get('@odata.nextLink')
    moved=0
    for mid in ids:
        if not apply: continue
        path='/me/messages/'+urllib.parse.quote(mid,safe='')
        current=graph.graph_request('GET',path,query={'$select':SELECT})
        if not eligible(current): continue
        result=graph.graph_request('POST',path+'/move',{'destinationId':DESTINATION})
        if not result.get('id'): raise RuntimeError('MOVE_NOT_VERIFIED')
        moved+=1
    return {'checked':checked,'matched':len(ids),'moved':moved}

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--apply',action='store_true')
    args=parser.parse_args()
    try: print(json.dumps(scan(args.apply)))
    except Exception:
        print('{"error":"MAIL_CLEANUP_FAILED"}')
        return 1
    return 0

if __name__=='__main__': raise SystemExit(main())

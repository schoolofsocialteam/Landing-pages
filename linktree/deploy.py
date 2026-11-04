"""Deploy linktree/site to Vercel production.

Auth: the Authorization header is added by the environment's network secret for api.vercel.com.
Usage: python3 linktree/deploy.py
"""
import hashlib, json, os, urllib.request

TEAM = 'team_EmAkDXog6gQJdgCTA5IDHNfD'
NAME = 'schoolofsocial-linkhub'
SITE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'site')
API = 'https://api.vercel.com'


def req(method, path, data=None, headers=None):
    r = urllib.request.Request(API + path, data=data, method=method, headers=headers or {})
    try:
        with urllib.request.urlopen(r) as res:
            body = res.read()
            return res.status, json.loads(body) if body else {}
    except urllib.error.HTTPError as e:
        return e.code, json.loads(e.read() or b'{}')


files = []
for root, _, names in os.walk(SITE):
    for n in sorted(names):
        p = os.path.join(root, n)
        b = open(p, 'rb').read()
        sha = hashlib.sha1(b).hexdigest()
        st, res = req('POST', '/v2/files?teamId=' + TEAM, b, {
            'Content-Type': 'application/octet-stream', 'x-vercel-digest': sha, 'Content-Length': str(len(b))})
        assert st in (200, 201), (n, st, res)
        files.append({'file': os.path.relpath(p, SITE), 'sha': sha, 'size': len(b)})
        print('uploaded', files[-1]['file'], len(b))

body = {'name': NAME, 'files': files, 'target': 'production',
        'projectSettings': {'framework': None, 'buildCommand': None, 'outputDirectory': None, 'installCommand': None}}
st, res = req('POST', '/v13/deployments?teamId=' + TEAM + '&skipAutoDetectionConfirmation=1',
              json.dumps(body).encode(), {'Content-Type': 'application/json'})
print(st, res.get('id'), res.get('url'), res.get('readyState'), res.get('error'))

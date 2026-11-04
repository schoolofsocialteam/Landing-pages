"""Build production files for Vercel from the page sources.

Usage: python3 linktree/build.py
Output: linktree/site/  (index.html, apply.html, og images, favicon, vercel.json)
"""
import os, re, shutil, json

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'site')
SITE = 'https://schoolofsocial.co'

PAGES = {
    'index.html': {
        'desc': 'Arturo Leduc, School of Social. Join the community, book a consultation, or apply to work with Arturo 1:1.',
        'og_title': 'Arturo Leduc / School of Social',
        'og_image': SITE + '/og.png',
        'url': SITE + '/',
    },
    'apply.html': {
        'desc': 'Apply to work directly with Arturo. 1:1 content coaching with limited spots.',
        'og_title': 'Apply to work directly with Arturo',
        'og_image': SITE + '/og-apply.png',
        'url': SITE + '/apply',
    },
}

BASE_CSS = ('<style>'
            '*,*::before,*::after{box-sizing:border-box}'
            ':root{padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}'
            'body{margin:0}img{max-width:100%}[hidden]{display:none!important}'
            '</style>')


def build(name, meta):
    src = open(os.path.join(HERE, name)).read()
    split = src.index('<div class="mockbar">')
    head, body = src[:split], src[split:]
    # drop mockup bar
    body = re.sub(r'<div class="mockbar">.*?</div>\n?', '', body, count=1, flags=re.S)
    # logo links home instead of a fake URL
    body = body.replace('<a class="wm" href="#" data-go="schoolofsocial.com" data-name="Website">',
                        '<a class="wm" href="/" aria-label="School of Social home">')
    assert 'data-go=' not in body.split('<script>')[0], name + ': fake link left in markup'
    tags = '\n'.join([
        '<meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">',
        '<meta name="description" content="%s">' % meta['desc'],
        '<meta name="theme-color" content="#FAF9F6">',
        '<link rel="icon" type="image/svg+xml" href="/favicon.svg">',
        '<link rel="apple-touch-icon" href="/apple-touch-icon.png">',
        '<meta property="og:type" content="website">',
        '<meta property="og:url" content="%s">' % meta['url'],
        '<meta property="og:title" content="%s">' % meta['og_title'],
        '<meta property="og:description" content="%s">' % meta['desc'],
        '<meta property="og:image" content="%s">' % meta['og_image'],
        '<meta property="og:image:width" content="1200">',
        '<meta property="og:image:height" content="630">',
        '<meta name="twitter:card" content="summary_large_image">',
        BASE_CSS,
    ])
    html = '<!doctype html>\n<html lang="en">\n<head>\n' + tags + '\n' + head + '</head>\n<body>\n' + body + '\n</body>\n</html>\n'
    open(os.path.join(OUT, name), 'w').write(html)


if __name__ == '__main__':
    os.makedirs(OUT, exist_ok=True)
    for n, m in PAGES.items():
        build(n, m)
    for f in ['favicon.svg', 'apple-touch-icon.png', 'og.png', 'og-apply.png']:
        p = os.path.join(HERE, 'assets', f)
        if os.path.exists(p):
            shutil.copy(p, os.path.join(OUT, f))
    json.dump({'cleanUrls': True, 'trailingSlash': False}, open(os.path.join(OUT, 'vercel.json'), 'w'), indent=2)
    print('built', sorted(os.listdir(OUT)))

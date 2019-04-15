import collections
import copy
import getpass
import json
import lxml
import lxml.etree
import os
import pathlib
import sys
import urllib
import zipfile

from confluence import client


chm_file = sys.argv[1]
if chm_file.lower().endswith('.chm'):
    chm_file = chm_file[:-4]

chm_short_name = os.path.splitext(os.path.split(chm_file)[-1])[0]
beckhoff_to_confluence_fn = f'{chm_short_name}.map.json'

extracted_path = pathlib.Path(sys.argv[2])

space_key = 'SBI'

source_extensions = {'.htm', '.html'}
source_by_id = {}
special_paths = {}
assets_by_path = {}


def rewrite_link_for_confluence(source_path, beckhoff_to_confluence):
    if not source_path:
        return ''
    elif source_path.startswith('http'):
        return source_path
    elif source_path.startswith('mailto'):
        return source_path

    if source_path in beckhoff_to_confluence['by_file']:
        confluence_id = beckhoff_to_confluence['by_file'][source_path]
        return f'/pages/viewpage.action?pageId={confluence_id}'

    return source_path


def parse_html(html_path, contents):
    tree = lxml.etree.fromstring(contents)
    metadata = collections.defaultdict(list)
    for md in tree.findall('.//meta', namespaces=tree.nsmap):
        md = dict(md.items())
        if 'name' in md and 'content' in md:
            metadata[md['name']].append(md['content'])

    metadata['parent_path'] = html_path

    return dict(metadata), tree


class HelpItem:
    def __init__(self, beckhoff_id, filename, *, parent=None, confluence_id=None):
        self.beckhoff_id = beckhoff_id
        self.confluence_id = confluence_id
        self.filename = filename
        self.metadata = {'filename': filename,
                         'beckhoff-id': self.beckhoff_id,
                         'chm-file': chm_file,
                         }

        with open(extracted_path / filename, 'rt', encoding='Windows-1252') as f:
            self.contents = f.read()

        self.tree = lxml.etree.fromstring(self.contents)
        try:
            title = get_title(self.tree)
        except Exception:
            title = self.beckhoff_id

        self.title = f'{title} ({beckhoff_id})'
        self.children = []
        self.parent = parent

    def __repr__(self):
        return (f'<HelpItem {self.beckhoff_id} ({self.confluence_id}) {self.title!r} '
                f'children={len(self.children)}>')


def get_order(path, fn='index.hhc'):
    with open(path / fn, 'rt', encoding='Windows-1252') as f:
        text = f.read()

    return [line.split('value="')[1].rstrip('>"')
            for line in text.splitlines() if 'Local' in line]


def find_by_path(path):
    path = str(path)
    return [s for s in source_by_id.values()
            if str(s['dest_path']).startswith(path)]


def get_id(doc):
    filename = os.path.split(doc)[-1]
    doc, ext = os.path.splitext(filename)
    return f'{chm_short_name}_{doc}'


def get_title(tree):
    return tree.findall('.//title', namespaces=tree.nsmap)[0].text


def wrap_html(html):
    return '''\
<ac:structured-macro ac:name="html">
  <ac:plain-text-body><![CDATA[{}]]></ac:plain-text-body>
</ac:structured-macro>
'''.format(html)


def build_page(item, dry_run=True):
    pg = c.get_content_by_id(
        content_id=item.confluence_id,
        expand=['extensions', 'space', 'body', 'body.view', 'history',
                'version'])

    if pg.space.key != space_key:
        raise ValueError(f'Unexpected space: {pg.space.key}')

    tree = copy.deepcopy(item.tree)
    images = tree.findall('.//img', namespaces=tree.nsmap)
    for img in images:
        src = img.get('src')
        if src:
            fn = extracted_path / img.get('src')
            fn = urllib.parse.unquote(str(fn))

            assert os.path.exists(fn)
            remote_fn = os.path.split(fn)[-1]
            img.set('src', f'/download/attachments/{item.confluence_id}/{remote_fn}')
            if dry_run:
                print('attach to', item.confluence_id, 'file', fn, 'remote=',
                      remote_fn)
            else:
                try:
                    c.add_attachment(content_id=item.confluence_id, file_path=fn,
                                     file_name=remote_fn)
                except client.ConfluenceError:
                    # c.update_attachment()
                    print('TODO', fn)

    links = (tree.findall('.//a', namespaces=tree.nsmap) +
             tree.findall('.//link', namespaces=tree.nsmap)
             )
    for link in links:
        href = link.get('href')
        if href:
            link.set('href',
                     rewrite_link_for_confluence(href, beckhoff_to_confluence))

    # Content properties are created in create_outline
    # c.create_content_property(pg.id, 'beckhoff', source_md['metadata'])

    new_content = wrap_html(lxml.etree.tostring(tree).decode('utf-8'))
    if dry_run:
        print('content', pg.id, new_content)
    else:
        c.update_content(
            content_id=pg.id,
            content_type=client.ContentType.PAGE,
            new_version=pg.version.number + 1,
            new_title=pg.title,  # source_md['metadata']['Title'],
            new_content=new_content,
        )


def create_outline(item):
    parent = item.parent
    parent_id = (parent.confluence_id if parent else None)
    if item.confluence_id is None:
        content = c.create_content(
            client.ContentType.PAGE, title=item.title, space_key=space_key,
            content='', parent_content_id=parent_id)
        item.confluence_id = content.id
        beckhoff_to_confluence['by_id'][item.beckhoff_id] = item.confluence_id
        beckhoff_to_confluence['by_file'][item.filename] = item.confluence_id

    try:
        c.create_content_property(item.confluence_id, 'beckhoff', item.metadata)
    except client.ConfluenceVersionConflict:
        ...

    for child in item.children:
        create_outline(child)

def walk(item):
    yield item
    for child in item.children:
        yield from walk(child)


def get_id_map(hier):
    return {item.beckhoff_id: item.confluence_id
            for top in hier
            for item in walk(top)
            }


def build_all(dry_run=True):
    for item in hier:
        build_page(item, dry_run=dry_run)


files = get_order(extracted_path)
try:
    with open(beckhoff_to_confluence_fn, 'rt') as f:
        beckhoff_to_confluence = json.load(f)
except Exception:
    beckhoff_to_confluence = {'by_file': {},
                              'by_id': {},
                              }

hier = [HelpItem(get_id(fn), filename=fn,
                 confluence_id=beckhoff_to_confluence['by_file'].get(fn))
        for fn in files]

# Make the first one the parent of all
root = hier[0]
root.children = hier[1:]
for child in root.children:
    child.parent = root

# import confluence.models.content
# c.delete_content(child.confluence_id, confluence.models.content.ContentStatus.CURRENT)

c = client.Confluence(
    'https://confluence.slac.stanford.edu',
    (getpass.getuser(), getpass.getpass())
)


c.__enter__()
# create_outline(root)
# build_all(dry_run=True)
# with open(beckhoff_to_confluence_fn, 'wt') as f:
#     json.dump(beckhoff_to_confluence, f)
#

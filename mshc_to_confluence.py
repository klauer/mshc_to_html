import collections
import copy
import getpass
import json
import lxml
import lxml.etree
import os
import pathlib
import sys
import zipfile

from confluence import client


output_path = pathlib.Path(sys.argv[1])
mshc_file = sys.argv[2]  #  'bkinfosys3_vs_100_en-us.mshc'
# space_root = 'https://confluence.slac.stanford.edu/display/SBI/'
space_key = 'SBI'

source_extensions = {'.htm', '.html'}
source_by_id = {}
special_paths = {}
assets_by_path = {}
SHARED_ATTACHMENT_ID = 245718672


def get_dest_path(source_path, relative_to=None):
    if not source_path:
        return ''
    elif source_path.startswith('http'):
        return source_path
    elif source_path.startswith('mailto'):
        return source_path

    if '?Id=' in source_path:
        _, doc_id = source_path.split('?Id=', 1)
        dest = source_by_id[doc_id]['dest_path']
    elif source_path in special_paths:
        dest = special_paths[source_path]
    else:
        # for path in strip_paths:
        #     if source_path.startswith(path):
        #         source_path = source_path[len(path):]

        dest = output_path / source_path.lstrip('/')

    if relative_to:
        return os.path.relpath(str(dest), relative_to)
    return str(dest)


def rewrite_link_for_confluence(source_path, beckhoff_to_confluence):
    if not source_path:
        return ''
    elif source_path.startswith('http'):
        return source_path
    elif source_path.startswith('mailto'):
        return source_path

    if '?Id=' in source_path:
        _, doc_id = source_path.split('?Id=', 1)
        confluence_id = beckhoff_to_confluence[doc_id]
        return f'/pages/viewpage.action?pageId={confluence_id}'
    elif source_path in special_paths:
        fn = os.path.split(source_path)[-1]
        # TODO hard-coded root id
        return f'/download/attachments/{SHARED_ATTACHMENT_ID}/{fn}'

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
    def __init__(self, beckhoff_id, info, *, parent=None, confluence_id=None):
        self.beckhoff_id = beckhoff_id
        self.confluence_id = None
        self.info = info
        self.metadata = get_md_for_confluence(info.get('metadata', {}))
        if 'Title' in self.metadata:
            self.title = self.metadata['Title']
        elif 'Description' in self.metadata and len(self.metadata['Description']) < 30:
            self.title = self.metadata['Description']
        else:
            self.title = str(beckhoff_id)

        self.title = f'{self.title} ({beckhoff_id})'
        self.children = []
        self.parent = parent


def build_index_hierarchy():
    items = list(sorted((id_, info['parent'])
                        for id_, info in source_by_id.items()))
    grouped_by_parent = collections.defaultdict(list)

    for id_, parent in items:
        grouped_by_parent[parent].append(id_)

    top_levels = [HelpItem(id_, source_by_id[id_])
                  for id_, parent in items
                  if parent not in source_by_id]

    def build(parent):
        for child_id in grouped_by_parent[parent.beckhoff_id]:
            child = HelpItem(child_id, source_by_id[child_id], parent=parent)
            parent.children.append(child)
            build(child)

    for top in top_levels:
        build(top)

    return top_levels


with zipfile.ZipFile(mshc_file, 'r') as zf:
    for finfo in zf.filelist:
        source_path = finfo.filename
        suffix = pathlib.Path(source_path).suffix
        dest_path = pathlib.Path(get_dest_path(source_path))

        with zf.open(finfo, 'r') as f:
            contents = f.read()

        if suffix in source_extensions:
            contents = contents.decode('utf-8')

            metadata, tree = parse_html(dest_path, contents)

            source_id = metadata['Microsoft.Help.Id'][0]
            source_by_id[source_id] = {
                'dest_path': dest_path,
                'tree': tree,
                'id': source_id,
                'parent': metadata.get('Microsoft.Help.TOCParent', [None])[0],
                'metadata': metadata,
            }

            continue

        if dest_path.parent == output_path:
            special_paths[source_path] = dest_path

        # os.makedirs(dest_path.parent, exist_ok=True)
        # with open(dest_path, 'wb') as df:
        #     df.write(contents)
        assets_by_path[dest_path] = contents


def find_by_path(path):
    path = str(path)
    return [s for s in source_by_id.values()
            if str(s['dest_path']).startswith(path)]


for source_id, info in source_by_id.items():
    tree = info['tree']

    parent_path = info['dest_path'].parent

    # images = tree.findall('.//img', namespaces=tree.nsmap)
    # for img in images:
    #     src = img.get('src')
    #     if src:
    #         img.set('src', str(get_dest_path(src, parent_path)))

    # links = (tree.findall('.//a', namespaces=tree.nsmap) +
    #          tree.findall('.//link', namespaces=tree.nsmap)
    #          )

    # for link in links:
    #     href = link.get('href')
    #     if href:
    #         new_href = str(get_dest_path(href, parent_path))
    #         link.set('href', new_href)

    # contents = lxml.etree.tostring(tree)
    # os.makedirs(parent_path, exist_ok=True)

    # with open(info['dest_path'], 'wb') as df:
    #     df.write(contents)
    # assets_by_path[info['dest_path']] = contents


hier = build_index_hierarchy()
# create_index(index)


def wrap_html(html):
    return '''\
<ac:structured-macro ac:name="html">
  <ac:plain-text-body><![CDATA[{}]]></ac:plain-text-body>
</ac:structured-macro>
'''.format(html)


def get_md_for_confluence(md):
    def get_value(value):
        if isinstance(value, list):
            if len(value) == 1:
                return get_value(value[0])
            return [get_value(v) for v in value]
        return str(value)
    return {key: get_value(value) for key, value in md.items()}


def build_page(beckhoff_to_confluence, confluence_id):
    confluence_to_beckhoff = dict((v, k) for k, v in beckhoff_to_confluence.items())
    beckhoff_id = confluence_to_beckhoff[confluence_id]
    source_md = source_by_id[beckhoff_id]

    pg = c.get_content_by_id(
        content_id=confluence_id,
        expand=['extensions', 'space', 'body', 'body.view', 'history',
                'version'])

    if pg.space.key != space_key:
        raise ValueError(f'Unexpected space: {pg.space.key}')

    tree = copy.deepcopy(source_md['tree'])
    images = tree.findall('.//img', namespaces=tree.nsmap)
    for img in images:
        src = img.get('src')
        if src:
            fn = os.path.split(img.get('src'))[-1]
            img.set('src', f'/download/attachments/{confluence_id}/{fn}')
            with open(fn, 'wb') as f:
                key = output_path / src.lower().lstrip('/')
                f.write(assets_by_path[key])

            try:
                c.add_attachment(content_id=confluence_id, file_path=fn,
                                 file_name=fn)
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
    # c.create_content_property(pg.id, 'beckhoff',
    #                           get_md_for_confluence(source_md['metadata']))

    c.update_content(
        content_id=pg.id,
        content_type=client.ContentType.PAGE,
        new_version=pg.version.number + 1,
        new_title=pg.title,  # source_md['metadata']['Title'],
        new_content=wrap_html(lxml.etree.tostring(tree).decode('utf-8')),
    )


def create_outline(item):
    parent = item.parent
    parent_id = (parent.confluence_id if parent else None)
    if item.confluence_id is None:
        content = c.create_content(
            client.ContentType.PAGE, title=item.title, space_key=space_key,
            content='', parent_content_id=parent_id)
        item.confluence_id = content.id

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


def build_all():
    for i, (beckhoff_id, confluence_id) in enumerate(beckhoff_to_confluence.items()):
        if (i % 1000) == 0:
            for j in range(20):
                print('----------------------', i)

        build_page(beckhoff_to_confluence, confluence_id=confluence_id)


c = client.Confluence(
    'https://confluence.slac.stanford.edu',
    (getpass.getuser(), getpass.getpass())
)


c.__enter__()

with open('beckhoff_to_confluence.json', 'rt') as f:
    beckhoff_to_confluence = json.load(f)

# source_md, = find_by_path('bkhf-confluence/tf8810_tc3_aes70/1033/index.html')
# build_page(245717810, source_md)
# c.update_content_property(
#     pg.id, 'beckhoff-page', beckhoff_id, pg.version.number + 2,
#     is_minor_edit=True, is_hidden_edit=True)

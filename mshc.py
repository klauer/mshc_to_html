import collections
import lxml
import lxml.etree
import os
import pathlib
import sys
import zipfile

output_path = pathlib.Path(sys.argv[1])
mshc_file = sys.argv[2]  #  'bkinfosys3_vs_100_en-us.mshc'

source_extensions = {'.htm', '.html'}
source_by_id = {}
special_paths = {}


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


def parse_html(html_path, contents):
    tree = lxml.etree.fromstring(contents)
    metadata = collections.defaultdict(list)
    for md in tree.findall('.//meta', namespaces=tree.nsmap):
        md = dict(md.items())
        if 'name' in md and 'content' in md:
            metadata[md['name']].append(md['content'])

    metadata['parent_path'] = html_path

    return dict(metadata), tree


def build_index_hierarchy():
    hierarchy = {}
    index = []

    items = list(sorted((id_, info['parent'])
                        for id_, info in source_by_id.items()))
    grouped_by_parent = collections.defaultdict(list)

    for id_, parent in items:
        grouped_by_parent[parent].append(id_)

    top_levels = [id_ for id_, parent in items
                  if parent not in source_by_id]

    def build(parents, parent_dict, depth):
        parent_id = parents[-1]
        index.append(parents)

        this_dict = {}
        parent_dict[parent_id] = this_dict
        for child_id in grouped_by_parent[parent_id]:
            build(parents + [child_id], this_dict, depth=depth + 1)

    for top in top_levels:
        build([top], hierarchy, depth=0)

    return hierarchy, index


def create_index(index_lines):
    indexed = []
    for ids in index_lines:
        info = source_by_id[ids[-1]]
        dest_path = info['dest_path']
        try:
            title = info['metadata']['Title'][0]
        except KeyError:
            title = str(dest_path)

        try:
            desc = info['metadata']['Description'][0]
        except KeyError:
            desc = ''
        else:
            if desc == title:
                desc = ''

        indexed.append((ids, title, desc, str(dest_path.relative_to(output_path))))

    with open(output_path / 'index.html', 'wt') as f:
        for ids, title, desc, dest_path in sorted(indexed):
            depth = len(ids)
            print(
                '&nbsp;' * depth,
                f'<a title={desc!r} href={dest_path!r}>{title}</a>',
                f'<br />', file=f)


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

        os.makedirs(dest_path.parent, exist_ok=True)
        with open(dest_path, 'wb') as df:
            df.write(contents)


for source_id, info in source_by_id.items():
    tree = info['tree']

    parent_path = info['dest_path'].parent

    images = tree.findall('.//img', namespaces=tree.nsmap)
    for img in images:
        src = img.get('src')
        if src:
            img.set('src', str(get_dest_path(src, parent_path)))

    links = (tree.findall('.//a', namespaces=tree.nsmap) +
             tree.findall('.//link', namespaces=tree.nsmap)
             )

    for link in links:
        href = link.get('href')
        if href:
            new_href = str(get_dest_path(href, parent_path))
            link.set('href', new_href)

    contents = lxml.etree.tostring(tree)
    os.makedirs(parent_path, exist_ok=True)

    with open(info['dest_path'], 'wb') as df:
        df.write(contents)


hier, index = build_index_hierarchy()
create_index(index)

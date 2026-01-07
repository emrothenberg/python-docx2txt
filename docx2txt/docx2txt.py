#! /usr/bin/env python

import argparse
import re
from typing import List, Literal, Union, cast, overload
import xml.etree.ElementTree as ET
import zipfile
import os
import sys


nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def process_args():
    parser = argparse.ArgumentParser(description='A pure python-based utility '
                                                 'to extract text and images '
                                                 'from docx files.')
    parser.add_argument("docx", help="path of the docx file")
    parser.add_argument('-s', '--split_pages',
                        help='Split text on page breaks')
    parser.add_argument('-i', '--img_dir', help='path of directory '
                                                'to extract images')

    args = parser.parse_args()

    if not os.path.exists(args.docx):
        print('File {} does not exist.'.format(args.docx))
        sys.exit(1)

    if args.img_dir is not None:
        if not os.path.exists(args.img_dir):
            try:
                os.makedirs(args.img_dir)
            except OSError:
                print("Unable to create img_dir {}".format(args.img_dir))
                sys.exit(1)
    return args


def qn(tag):
    """
    Stands for 'qualified name', a utility function to turn a namespace
    prefixed tag name into a Clark-notation qualified tag name for lxml. For
    example, ``qn('p:cSld')`` returns ``'{http://schemas.../main}cSld'``.
    Source: https://github.com/python-openxml/python-docx/
    """
    prefix, tagroot = tag.split(':')
    uri = nsmap[prefix]
    return '{{{}}}{}'.format(uri, tagroot)


@overload
def xml2text(xml, split_pages: Literal[True]) -> List[str]: ...


@overload
def xml2text(xml, split_pages: Literal[False]) -> str: ...


def xml2text(xml, split_pages: bool):
    """
    A string representing the textual content of this run, with content
    child elements like ``<w:tab/>`` translated to their Python
    equivalent.
    Adapted from: https://github.com/python-openxml/python-docx/
    """
    text = u''
    texts = []

    root = ET.fromstring(xml)
    for child in root.iter():
        if child.tag == qn('w:t'):
            t_text = child.text
            text += t_text if t_text is not None else ''
        elif child.tag == qn('w:tab'):
            text += '\t'
        elif split_pages and child.tag == qn('w:br') and list(child.attrib.values())[0] == "page":
            texts.append(text)
            text = u''
        elif child.tag in (qn('w:br'), qn('w:cr')):
            text += '\n'
        elif child.tag == qn("w:p"):
            text += '\n\n'
    return texts if split_pages else text


def strip_list(lst: list):
    while lst and not lst[-1]:
        lst.pop()

    while lst and not lst[0]:
        lst.pop(0)

    return lst


@overload
def process(docx, split_pages: Literal[True], img_dir=None) -> List[str]: ...


@overload
def process(docx, split_pages: Literal[False], img_dir=None) -> str: ...


def process(docx, split_pages=False, img_dir=None):
    text: Union[list[str], str] = [] if split_pages else ""

    # unzip the docx in memory
    zipf = zipfile.ZipFile(docx)
    filelist = zipf.namelist()

    # get header text
    # there can be 3 header files in the zip
    header_xmls = 'word/header[0-9]*.xml'
    for fname in filelist:
        if re.match(header_xmls, fname):
            if split_pages:
                text = cast(list[str], text)
                text.extend(xml2text(zipf.read(fname), split_pages))
            else:
                text = cast(str, text)
                text += xml2text(zipf.read(fname), split_pages)

    # get main text
    doc_xml = 'word/document.xml'
    if split_pages:
        text = cast(list[str], text)
        text.extend(xml2text(zipf.read(doc_xml), split_pages))
    else:
        text = cast(str, text)
        text += xml2text(zipf.read(doc_xml), split_pages)

    # get footer text
    # there can be 3 footer files in the zip
    footer_xmls = 'word/footer[0-9]*.xml'
    for fname in filelist:
        if re.match(footer_xmls, fname):
            if split_pages:
                text = cast(list[str], text)
                text.extend(xml2text(zipf.read(fname), split_pages))
            else:
                text = cast(str, text)
                text += xml2text(zipf.read(fname), split_pages)

    if img_dir is not None:
        # extract images
        for fname in filelist:
            _, extension = os.path.splitext(fname)
            if extension in [".jpg", ".jpeg", ".png", ".bmp"]:
                dst_fname = os.path.join(img_dir, os.path.basename(fname))
                with open(dst_fname, "wb") as dst_f:
                    dst_f.write(zipf.read(fname))

    zipf.close()
    return [t.strip() for t in strip_list(cast(list[str], text))] if split_pages else cast(str, text).strip()


if __name__ == '__main__':
    args = process_args()
    text = process(args.docx, args.split_pages, args.img_dir)
    sys.stdout.write(text.encode('utf-8'))

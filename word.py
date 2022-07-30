# -*- coding: UTF-8 -*-
import fnmatch
import re
import os
import time

from loguru import logger
from win32com import client as wc


def check(filePath):
    """Check word file and get words number"""
    encodes = ['gbk', 'utf-8']
    logger.debug('Checking encoding...')
    for encode in encodes:
        _word2txt(filePath, encode)

    output_path = os.path.abspath('./tmp_utf-8.txt')
    words_number = _calculate_words(output_path)

    return words_number


def _word2txt(filePath, encoding):
    output_path = os.path.abspath(f'./tmp_{encoding}.txt')

    wordappp = wc.Dispatch('Word.Application')

    doc = wordappp.Documents.Open(filePath, encoding)

    doc.SaveAs(output_path, 4, Encoding=encoding)  # 4 means save as text
    doc.Close()

    wordappp.Quit()
    time.sleep(2)


def _calculate_words(txt_path):
    """Calculate words number"""
    with open(txt_path, 'r', encoding='gbk') as f:
        content = f.read()
        cjkReg = re.compile(u'[\u1100-\u2999\u3001-\uFFFD]+?')
        trimedCJK = cjkReg.sub(' a ', content, 0)   # replace the CJK with the word a
        # words = trimedCJK.split()
        return len(trimedCJK.split())

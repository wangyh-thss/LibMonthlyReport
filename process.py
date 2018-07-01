# encoding=utf-8
import re
import os
import xlrd
import xlwt


PAPER_TYPE_ATTR = 'type'
DEPARTMENT_MAP_FILE = 'map.xls'
ADDRESS_NUM = 'addr_num'
ADDRESS_INFO = 'addresses'
tsinghua_names = ['tsinghua', 'tsing hua', 'qinghua', 'qing hua']
except_keywords = ['natl tsing', '100871', 'peking tsinghua ctr', # 'peking tsinghua ctr life sci'
                   'peking univ tsinghua univ natl inst biol sci join, beijing 102206',
                   'peking univ, peking univ tsinghua univ natl inst biol sci join']
used_attrs = {'au', 'ti', 'so', 'dt', 'py', 'ut', 'la', 'c1', 'rp', 'em', 'tc', 'bn', 'sn', 'ei'}
output_attrs = [PAPER_TYPE_ATTR, 'au', 'ti', 'so', 'dt', 'rp', 'em', 'tc', 'py', 'ut', ADDRESS_NUM]
department_map = None


def get_except_keywords():
    global except_keywords
    return '\r\n'.join(except_keywords)


def set_except_keywords(s):
    global except_keywords
    except_keywords = map(lambda x: x.strip().lower(), s.splitlines())


def read_file(filename):
    global used_attrs
    (_, temp_filename) = os.path.split(filename)
    if temp_filename.startswith('hot'):
        paper_type = 'hot'
    else:
        paper_type = 'highly'
    with open(filename, 'rU') as f:
        lines = f.readlines()
    header = map(lambda x: x.strip().lower(), lines[0].split('\t'))
    data_length = len(lines)
    attr_length = len(header)
    papers = dict()
    for i in xrange(1, data_length):
        paper = dict()
        values = map(str.strip, lines[i].split('\t'))
        for attr_index in xrange(attr_length):
            if header[attr_index] in used_attrs:
                paper[header[attr_index]] = values[attr_index]
        paper[PAPER_TYPE_ATTR] = paper_type
        papers[paper['ut']] = paper
    return papers


def read_department_map(filename):
    department_map = dict()
    xls_file = xlrd.open_workbook(filename)
    sheet = xls_file.sheet_by_index(0)
    for row_index in xrange(1, sheet.nrows):
        en = sheet.cell(rowx=row_index, colx=0).value.encode('utf-8')
        ch = sheet.cell(rowx=row_index, colx=1).value.encode('utf-8')
        department_map[en] = ch
    return department_map


def merge_paper_set(paper_set1, paper_set2):
    for wos, paper in paper_set2.items():
        if wos not in paper_set1:
            paper_set1[wos] = paper
        elif paper[PAPER_TYPE_ATTR] not in paper_set1[wos][PAPER_TYPE_ATTR]:
            paper_set1[wos][PAPER_TYPE_ATTR] += '&' + paper[PAPER_TYPE_ATTR]
    return paper_set1


def judge_tsinghua_error(addr):
    global except_keywords
    addr = addr.lower()
    for error_key in except_keywords:
        if error_key in addr:
            return True
    return False


def is_tsinghua_addr(addr):
    tsinghua = False
    addr = addr.lower()
    for tsinghua_name in tsinghua_names:
        if tsinghua_name in addr:
            tsinghua = True
    if not tsinghua:
        return False
    if judge_tsinghua_error(addr):
        return False
    return True


def find_tsinghua_address(address_str, auth_str=None):
    global tsinghua_names
    addresses = re.findall(r'\[(.+?)\]([^\[]*)', address_str)
    tsinghua_addr = list()
    if len(addresses) > 0:
        order = 0
        for auth, addr in addresses:
            for addr_single in addr.split(';'):
                addr_single = addr_single.strip()
                if not addr_single:
                    continue
                order += 1
                if is_tsinghua_addr(addr_single):
                    tsinghua_addr.append((auth, addr_single, order, translate_address(addr_single)))
    else:
        addresses = map(str.strip, address_str.split(';'))
        order = 0
        for addr in addresses:
            if not addr.strip():
                continue
            order += 1
            if is_tsinghua_addr(addr):
                tsinghua_addr.append((auth_str, addr, order, translate_address(addr)))
    return tsinghua_addr, order


def translate_address(addr):
    global department_map
    addr_split = addr.split(',')
    length = len(addr_split)
    dep = None
    if length == 3:
        dep = addr_split[0]
    elif length == 4:
        if 'Shenzhen' in addr:
            dep = ','.join(addr_split[:3])
        else:
            dep = ','.join(addr_split[:2])
    elif length == 5:
        dep = ','.join(addr_split[:3])
    elif length == 6:
        dep = ','.join(addr_split[:4])
    return department_map.get(dep, '')


def write_result(filename, papers):
    global output_attrs
    if not filename:
        return
    content = '\t'.join(output_attrs)
    content += '\tc1_au\tc1_addr\tc1_order\tdepartment\tinstitute\r\n'
    for _, info in papers.items():
        info_content = ''
        for attr in output_attrs:
            info_content += str(info.get(attr, '')) + '\t'
        for (auth, addr, order, dep) in info[ADDRESS_INFO]:
            content += '%s\t%s\t%s\t%s\t%s\t%s\r\n' % (info_content, auth, addr, order, dep, '')
    with open(filename, 'w') as f:
        f.write(content)


def write_xlsx_result(filename, papers):
    global output_attrs
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('sheet1')
    for i, attr in enumerate(output_attrs + ['c1_au', 'c1_addr', 'c1_order', 'department', 'institute']):
        worksheet.write(0, i, label=attr)
    line = 1
    for _, info in papers.items():
        for values in info[ADDRESS_INFO]:
            i = 0
            for attr in output_attrs:
                worksheet.write(line, i, label=str(info.get(attr, ''))[:32767])
                i += 1
            for value in values:
                worksheet.write(line, i, label=value)
                i += 1
            line += 1
    workbook.save(filename)


def merge_paper_addr_by_department(addresses):
    keys = dict()
    for values in addresses:
        auth, addr_single, order, department = values
        key = '%s&%s' % ('', department)
        if key not in keys:
            keys[key] = values
        else:
            _, _, order_existing, _ = keys[key]
            keys[key] = (auth, addr_single, '%s,%s' % (order_existing, order), department)
    return keys.values()



def write_deduplicate_resutl(filename, papers):
    global output_attrs
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('sheet1')
    for i, attr in enumerate(output_attrs + ['c1_au', 'c1_addr', 'c1_order', 'department', 'institute']):
        worksheet.write(0, i, label=attr)
    line = 1
    for _, info in papers.items():
        for values in merge_paper_addr_by_department(info[ADDRESS_INFO]):
            i = 0
            for attr in output_attrs:
                worksheet.write(line, i, label=str(info.get(attr, ''))[:32767])
                i += 1
            for value in values:
                worksheet.write(line, i, label=value)
                i += 1
            line += 1
    workbook.save(filename)


def main(filenames, output_filename):
    global department_map, DEPARTMENT_MAP_FILE
    department_map = read_department_map(DEPARTMENT_MAP_FILE)
    papers = dict()
    for filename in filenames:
        papers = merge_paper_set(papers, read_file(filename))
    for _, info in papers.items():
        info[ADDRESS_INFO], info[ADDRESS_NUM] = find_tsinghua_address(info['c1'], info['au'])
    tsv_filename = output_filename + '.tsv'
    xls_filename = output_filename + '.xls'
    deduplicate_filename = u'%s_按学院去重.xls' % output_filename
    write_result(tsv_filename, papers)
    write_xlsx_result(xls_filename, papers)
    write_deduplicate_resutl(deduplicate_filename, papers)
    return tsv_filename, xls_filename, deduplicate_filename


if __name__ == '__main__':
    # files = ['../highly-1311.txt', '../hot-56.txt']
    file_dir = '../ESI2018年5月更新/'
    files = os.listdir(file_dir)
    files = filter(lambda x: x.startswith('highly') or x.startswith('hot'), files)
    files = map(lambda x: file_dir + x, files)
    # files = ['../lib/highly-PLANT & ANIMAL SCIENCE_389.txt']
    main(files, 'result')


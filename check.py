# encoding=utf-8


def read_file(filename):
    with open(filename, 'rU') as f:
        lines = f.readlines()
    header = map(lambda x: x.strip().lower(), lines[0].split('\t'))
    wos_index = header.index('ut')
    data_length = len(lines)
    papers = dict()
    for i in xrange(1, data_length):
        values = map(str.strip, lines[i].split('\t'))
        wos = values[wos_index]
        if wos not in papers:
            papers[wos] = 0
        papers[wos] += 1
    return papers


if __name__ == '__main__':
    truth = read_file('../ESI2018年5月更新/zhao.txt')
    test = read_file('result.tsv')
    wos_set = set(truth.keys() + test.keys())
    content = 'length: ' + str(len(wos_set)) + '\r\n'
    for wos in wos_set:
        truth_num = truth.get(wos, 0)
        test_num = test.get(wos, 0)
        if truth_num != test_num:
            content += '%s\t%s\t%s\r\n' % (wos, truth_num, test_num)
    print content


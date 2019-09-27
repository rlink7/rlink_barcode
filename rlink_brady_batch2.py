#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Copyright (c) 2019 CEA
#
# This software is governed by the CeCILL license under French law and
# abiding by the rules of distribution of free software. You can use,
# modify and/ or redistribute the software under the terms of the CeCILL
# license as circulated by CEA, CNRS and INRIA at the following URL
# "http://www.cecill.info".
#
# As a counterpart to the access to the source code and rights to copy,
# modify and redistribute granted by the license, users are provided only
# with a limited warranty and the software's author, the holder of the
# economic rights, and the successive licensors have only limited
# liability.
#
# In this respect, the user's attention is drawn to the risks associated
# with loading, using, modifying and/or developing or reproducing the
# software by the user in light of its specific status of free software,
# that may mean that it is complicated to manipulate, and that also
# therefore means that it is reserved for developers and experienced
# professionals having in-depth computer knowledge. Users are therefore
# encouraged to load and test the software's suitability as regards their
# requirements in conditions enabling the security of their systems and/or
# data to be ensured and, more generally, to use and operate it in the
# same conditions as regards security.
#
# The fact that you are presently reading this means that you have had
# knowledge of the CeCILL license and that you accept its terms.


import pandas
from rlink_barcode import LABELS

TIMEPOINTS = (
    'M00',  # before Li initiation
    'M03',  # 3 months after Li initiation
)

INPUT = 'R-LiNK_patients_batch2_2019-09-16.xlsx'
OUTPUT = 'R-LiNK_labels_batch_2_2019-09-16.xlsx'


def main():
    converters = {
        'PSC1': str,
    }
    patients = pandas.read_excel(INPUT, converters=converters)

    labels = []
    for timepoint in TIMEPOINTS:
        for patient in patients['Patient']:
            barcode = patient + '-' + timepoint
            tmp = []
            for label in LABELS:
                tmp.append((barcode, label))
            if len(tmp) % 2:  # print complete rows of 2 labels
                tmp.append((barcode, ''))
            # the Brady BBP12 printer prints rows bottom to top
            # revert rows of 2 labels to print in "natural" order
            len_tmp = len(tmp)
            mid_tmp = len_tmp // 2
            lag_tmp = len_tmp -1
            for i in range(0, mid_tmp, 2):
                tmp[i], tmp[lag_tmp-i-1] = tmp[lag_tmp-i-1], tmp[i]
                tmp[i+1], tmp[lag_tmp-i] = tmp[lag_tmp-i], tmp[i+1]
            labels.extend(tmp)

    barcodes, types = zip(*labels)
    data = {
        'Barcode': barcodes,
        'Type': types,
    }
    df = pandas.DataFrame(data, columns=['Barcode', 'Type'])
    df.to_excel(OUTPUT, index=False)


if __name__ == '__main__':
    main()

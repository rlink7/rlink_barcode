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
)

INPUT = 'R-LiNK_patients_pilot_2019-04-04.xlsx'
OUTPUT = 'R-LiNK_labels_pilot_2019-04-04.xlsx'


def main():
    converters = {
        'PSC1': str,
    }
    patients = pandas.read_excel(INPUT, converters=converters)

    contents = []
    for timepoint in TIMEPOINTS:
        for patient in patients['Patient']:
            for label in LABELS:
                barcode = patient + '-' + timepoint
                contents.append((barcode, label))
            if len(LABELS) % 2:  # skip a label
                barcode = '-' * (6 + 1 + 3)
                contents.append((barcode, ''))

    barcodes, types = zip(*contents)
    data = {
        'Barcode': barcodes,
        'Type': types,
    }
    df = pandas.DataFrame(data, columns=['Barcode', 'Type'])
    df.to_excel(OUTPUT, index=False)


if __name__ == '__main__':
    main()

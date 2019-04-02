# Generation of patient sheets and labels

The idea is to assist and automate association of patient codes with
actual patients as much as possible. Barcodes are provided to avoid
manual input errors.

## Lists of patient pseudonym codes in Microsoft Excel

We have created multiple Excel tables with patient pseudonym codes:

|       | Date       | File name                               | Description            |
|:------|:----------:|:----------------------------------------|:-----------------------|
| Pilot | 2019-04-03 | _R-LiNK_barcodes_pilot_2019-04-03.xlsx_ | Test codes 998 and 999 |

## Generate barcodes in Microsoft Word

We have prepared Word templates for patient inclusion sheets and packaging labels:

| File name            | Description                                        |
|:---------------------|:---------------------------------------------------|
| _R-LiNK_sheets.docx_ | 1 patient inclusion sheet and 2 MRI session sheets |
| _R-LiNK_labels.docx_ | labels for packaging                               |

They contain
[mail merge fields](https://support.office.com/en-us/article/mail-merge-insert-merge-field-ad4a6f9b-c590-471e-b432-7d9cfff34890),
that are associated to the _Patient_ column of the above Excel table.

To print Word documents:
1. _Publipostage_
2. _Démarrer la fusion et le publipostage_ | _Lettres_
3. _Terminer & fusionner_ | _Imprimer les documents..._

Here are a few links of interest:
* [How to Print Barcodes With Excel and Word](https://www.clearlyinventory.com/how-to-print-barcodes-with-excel-and-word)
* [`DISPLAYBARCODE` and `MERGEBARCODE`: How to Insert or Mail Merge Barcodes](https://hubpages.com/technology/Mail-Mergeable-Barcodes-in-Microsoft-Word-2013-aka-Bar-Codes)
* [`DISPLAYBARCODE`](https://docs.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/cbc893c0-9683-416d-84c6-407a92451c19)
* [`MERGEBARCODE`](https://docs.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/cc4b13c2-c09b-4545-a6ae-4509d943233e)
* [`MERGEFIELD`](https://support.office.com/en-us/article/field-codes-mergefield-field-7a6d24a1-68a6-4b05-8359-1dc087daf4e6)

## Generate barcodes in LabelMark 6 for Brady labels

Here are a few links of interest:
* [How to Import Data from Excel to LabelMark 6 Software](https://www.youtube.com/watch?v=ISnkwf5efmg)

# Generation of patient sheets and labels

The idea is to assist and automate association of patient codes with
actual patients as much as possible. Barcodes are provided to avoid
manual input errors.

## Lists of patient pseudonym codes in Microsoft Excel

We have created multiple Excel tables with patient pseudonym codes and label contents:

|              | File                                                                                                                                                        | Description                                                               |
|:-------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------|:--------------------------------------------------------------------------|
| Pilot        | [R-LiNK_patients_pilot_2019-04-04.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/pilot_2019-04-08/R-LiNK_patients_pilot_2019-04-04.xlsx)    | Test codes for patients 998 and 999 in centre 01 to 16                    |
| Pilot        | [R-LiNK_labels_pilot_2019-04-04.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/pilot_2019-04-08/R-LiNK_labels_pilot_2019-04-04.xlsx)        | Labels generated for the above test patients - M00 time point only        |
| Batch&nbsp;1 | [R-LiNK_patients_batch1_2019-09-12.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch1_2019-09-13/R-LiNK_patients_batch1_2019-09-12.xlsx) | Codes for patients 001 to 020 in centres 01 to 03                         |
| Batch&nbsp;1 | [R-LiNK_labels_batch1_2019-09-12.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch1_2019-09-13/R-LiNK_labels_batch1_2019-09-12.xlsx)     | Labels generated for the above patients - time points M00 and M03         |
| Batch&nbsp;2 | [R-LiNK_patients_batch2_2019-09-16.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch2_2019-09-27/R-LiNK_patients_batch2_2019-09-16.xlsx) | Codes for patients 001 to 020 in centres 04-07, 09-10, 15                 |
| Batch&nbsp;2 | [R-LiNK_labels_batch2_2019-09-16.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch2_2019-09-27/R-LiNK_labels_batch2_2019-09-16.xlsx)     | Labels generated for the above patients - time points M00 and M03         |
| Batch&nbsp;3 | [R-LiNK_patients_batch3_2019-09-30.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch3_2019-09-30/R-LiNK_patients_batch3_2019-09-30.xlsx) | Codes for patients 001 to 020 in centres 08, 11-14, 16                    |
| Batch&nbsp;3 | [R-LiNK_labels_batch3_2019-09-30.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch3_2019-09-30/R-LiNK_labels_batch3_2019-09-30.xlsx)     | Labels generated for the above patients - time points M00 and M03         |
| Batch&nbsp;4 | [R-LiNK_patients_batch4_2020-10-26.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch4_2020-10-26/R-LiNK_patients_batch4_2020-10-26.xlsx) | Codes for 10 patients from all centres - Covid-19 crisis delay            |
| Batch&nbsp;4 | [R-LiNK_labels_batch4_2020-10-26.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch4_2020-10-26/R-LiNK_labels_batch4_2020-10-26.xlsx)     | Labels generated for the above patients - time points M00 and M03         |
| Batch&nbsp;5 | [R-LiNK_patients_batch5_2021-01-18.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch5_2021-01-18/R-LiNK_patients_batch5_2021-01-18.xlsx) | Codes for the remaining patients from all centres - Covid-19 crisis delay |
| Batch&nbsp;5 | [R-LiNK_labels_batch5_2021-01-18.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch5_2021-01-18/R-LiNK_labels_batch5_2021-01-18.xlsx)     | Labels generated for the above patients - time points M00 and M03         |
| Batch&nbsp;6 | [R-LiNK_patients_batch6_2021-05-20.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch6_2021-05-20/R-LiNK_patients_batch6_2021-05-20.xlsx) | Labels to re-print - some time points only                                |
| Batch&nbsp;6 | [R-LiNK_labels_batch6_2021-05-20.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch6_2021-05-20/R-LiNK_labels_batch6_2021-05-20.xlsx)     | Labels to re-print - some time points only                                |
| Batch&nbsp;7 | [R-LiNK_patients_batch7_2021-12-13.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch7_2021-12-13/R-LiNK_patients_batch7_2021-12-13.xlsx) | Codes for new patients 021 to 025 in centre 15                            |
| Batch&nbsp;7 | [R-LiNK_labels_batch7_2021-12-13.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch7_2021-12-13/R-LiNK_labels_batch7_2021-12-13.xlsx)     | Labels generated for the above patients - time points M00 and M03         |
| Batch&nbsp;8 | [R-LiNK_patients_batch8_2022-01-18.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch8_2022-01-18/R-LiNK_patients_batch8_2022-01-18.xlsx) | Re-print patients 001 to 005 in centre 12                                 |
| Batch&nbsp;8 | [R-LiNK_labels_batch8_2022-01-18.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch8_2022-01-18/R-LiNK_labels_batch8_2022-01-18.xlsx)     | Re-print labels for the above patients - time points M00 and M03          |
| Batch&nbsp;9 | [R-LiNK_patients_batch9_2022-03-11.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch9_2022-03-11/R-LiNK_patients_batch9_2022-03-11.xlsx) | Codes for new patients 026 to 030 in centre 15                            |
| Batch&nbsp;9 | [R-LiNK_labels_batch9_2022-03-11.xlsx](https://github.com/rlink7/rlink_barcode/blob/master/data/batch9_2022-03-11/R-LiNK_labels_batch9_2022-03-11.xlsx)     | Labels generated for the above patients - time points M00 and M03         |

## Generate barcodes in Microsoft Word

We have prepared Word templates for patient inclusion sheets and packaging labels:

| File                                                                                              | Description                                        |
|:--------------------------------------------------------------------------------------------------|:---------------------------------------------------|
| [R-LiNK_sheets.docx](https://github.com/rlink7/rlink_barcode/blob/master/data/R-LiNK_sheets.docx) | 1 patient inclusion sheet and 2 MRI session sheets |
| [R-LiNK_labels.docx](https://github.com/rlink7/rlink_barcode/blob/master/data/R-LiNK_labels.docx) | labels for packaging                               |

They contain [mail merge fields](https://support.office.com/en-us/article/mail-merge-insert-merge-field-ad4a6f9b-c590-471e-b432-7d9cfff34890) that can be associated to the _Patient_ column of the above Excel tables.

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

We have prepared a model to generate labels for biolgical samples:

| File                                                                                                                                               | Description                              |
|:---------------------------------------------------------------------------------------------------------------------------------------------------|:-----------------------------------------|
| [R-LiNK_labels_pilot_2019-04-08.l6t](https://github.com/rlink7/rlink_barcode/blob/master/data/pilot_2019-04-08/R-LiNK_labels_pilot_2019-04-08.l6t) | model of labels used for pilot           |
| [R-LiNK_labels.l6t](https://github.com/rlink7/rlink_barcode/blob/master/data/R-LiNK_labels.l6t)                                                    | model of labels used for batches 1, 2, 3 |

This model can be associated to the _Barcode_ and _Type_ columns of the above Excel tables.

Here are a few links of interest:
* [How to Import Excel Data into an Existing Template in LabelMark 6](https://support.bradyid.com/s/article/How-to-Import-Excel-Data-into-an-Existing-Template-in-LabelMark-6)

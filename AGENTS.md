# Working on the postal tool

The agreed product requirements and implementation stages are in PROJECT_PLAN_AR.md.

The daily workflow extracts the incoming number and date from the incoming stamp,
a short subject grounded in the letter, both the sending and addressed entities,
and receipt date/time generated at successful
save. Generate a unique serial automatically. QR is optional and must not override
an incoming number read from the stamp. An unrelated issue date is not the stamp date.
The user made incoming number and stamp date optional at save on 2026-10-11.
Leave missing values blank; do not require stamp-review confirmation for empty fields.
When the incoming number is absent, name the Word document and archive using the
automatically generated unique serial. The subject remains required.
The external entity field was removed on 2026-10-11. Do not extract it as a record
field, display it in the UI, or store it in Excel or Word metadata. Keep the
sending entity and addressed entity and preserve the full original OCR text.

Each successful save creates one Excel row and a Word document with the entire OCR
text, named for the incoming number. Scanner documents also retain their original
image with excessive blank scanner padding removed. The user explicitly requested
one archived image only, without a second full-length original copy. Keep all
document content and do not distort its proportions. Uploaded documents do not require an additional original copy. Never delete
the employee's source files or replace an existing archive.

Never run write tests against البوستة.xlsx or the production scans directory. Use
temporary workbooks and explicit POSTAL_SCANS, POSTAL_INBOX, and POSTAL_BACKUPS paths.
Keep documents, workbooks, private evaluation material, models, and snapshots out of
Git. Do not upload business documents to external services during development.

Useful checks: python -m unittest discover -s التطوير/اختبارات -p test_record_workflow.py -v;
python التطوير/اختبارات/test_evidence_contract.py; python التطوير/اختبارات/test_samples.py;
python التطوير/اختبارات/test_word_format.py. Legacy development scripts were moved
outside the project during the authorized fresh reset. Existing sample expectations cover
only three reviewed letters. Passing them is not a general accuracy guarantee.

Verify both stamp fields separately and report their exact match rates. OCR scores
are engine signals, not statistical probabilities. A small known sample is a
development regression set, not an independent acceptance set. Physical scanner
and clean installation checks require separate evidence.

The user requested a fresh start on 2026-10-11. Active registries and data folders
were reset after archiving all old records outside this project. Keep the working
folder free of sample scans and debug output. For old regression scans, explicitly
set POSTAL_SAMPLE_ROOT to the external archive and POSTAL_EVALUATION_DIR to a
non-production output folder. Never interpret skipped samples as passing OCR.

Presentation: receipt date is the first Excel column. The only extracted record date is the incoming stamp date. The user removed
the letter issue date on 2026-10-11; do not extract, display, store or recreate
that field in registries or Word metadata. Preserve the full original OCR text. Manual follow-up columns use a distinct pale fill.
Word contains a metadata table on the first page and full OCR text starting on
the following page, with readable Arabic type and no developer evidence report.

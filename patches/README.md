# Local QA pull fixes

`Skyline_pull.py` lives in the sibling `PB/Quality` folder, outside this Git
repository, and contains local credentials. Do not copy that full script here.

`skyline_transport_retry.patch` records the credential-free November 3, 2026
fix already applied to the local Skyline script. Weekly API requests retry
connection resets, interrupted bodies, timeouts, invalid JSON, and transient
HTTP errors up to four times, with delays of 3, 6, and 9 seconds. Permanent
HTTP failures stop immediately; exhausted retries raise and block publication.
The annual date range and safe RAW workbook replacement remain unchanged.

`data_carz_transport_retry.patch` and `kelowna_transport_retry.patch` apply
the same handling to Data Carz and Kelowna, which failed with the same
connection reset during the recovery run. They also add the `time` import
needed for backoff. All three scripts are covered by the transport tests.

`ollies_transport_retry.patch` adds the same proven handling for Ollies after
the November 4 refresh exposed the identical interrupted-response failure.

On another PC, apply this patch from its `PB` folder using
`git apply --unidiff-zero Masterlist/patches/skyline_transport_retry.patch`
(or merge the same hunks manually). Run from `Masterlist` afterward:

```powershell
py -3 -m unittest discover -s tests -v
py -3 -m py_compile ..\Quality\Skyline\Skyline_pull.py
```

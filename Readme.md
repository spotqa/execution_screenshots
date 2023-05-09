# Virtuoso execution step report

## Requirements
- Python >= 3.11
- pip >= 23.1.2

### Install python packages
```console
pip install -r requirements.txt
```

## Run
```console
Usage:  python main.py OPTIONS

Virtuoso test execution report self extractor.

Options:
        -t, --token uuid    Virtuoso token
        -i, --id int        Virtuoso execution id
        -e, --env string    Environment [OPTIONAL] (default = "api-app2.virtuoso.qa")>
        -b, --block         If present layout images are displayed vertical blocks [OPTIONAL] (default Side-by-side horizontal blocks)

Copyright SPOTQA LTD 2023 | <support@virtuoso.qa>
```

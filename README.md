# Google Drive Permission Search

Search/List your Google Drive's Documents.

Requirements:
- Ruby 2.0.0 or higher

# Installation

This script is originated by google-api-ruby-client-samples.

https://github.com/google/google-api-ruby-client-samples/tree/master/drive

- Get your client_secrets.json file from Google API Console
- Save your client_secrets.json to project root

For installation:

```
bundle
```

# Usage

```
bundle exec ./google-drive-permission-search.rb --verbose --type excel
```

# Options

- -type: "excel" or "tsv" is available (default: excel)
- -f FILE_NAME: specify output file name (default: result.(xls|tsv)])
- -v, --verbose: output debug messages
- --only-includes TEXT: output data which includes "TEXT" string only

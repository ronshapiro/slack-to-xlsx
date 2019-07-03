# Slack to .xlsx

A quick-and-dirty tool to convert a slack archive .zip file into an .xlsx file.

1. Download your archive at: https://your-workspace.slack.com/services/export
2. Run this in a terminal:

```sh
$ pip install XlsxWriter
$ python slack_to_xlsx.py <path/to/slack_archive.zip>
```

TODO(ronshapiro): Add a [Google Docs](https://developers.google.com/sheets/api/quickstart/python) frontend

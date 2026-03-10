# Contributing

Contributions are welcome.

## Setup

```bash
git clone https://github.com/mediatoolbox-org/slidemaker.git
cd slidemaker
mamba env create --file conda/dev.yaml
conda activate slidemaker
poetry config virtualenvs.create false
poetry install --with dev
```

## Local Checks

```bash
makim tests.unit
makim tests.linter
makim docs.build
```

## Pull Requests

- keep changes focused
- include tests for behavior changes
- update docs when user-facing behavior changes
- keep Python 3.10+ compatibility

## Release Model

This repository uses semantic-release. Follow conventional commit style in PR
titles/messages.

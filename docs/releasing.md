# Releasing

PyAccessKit is published to [PyPI](https://pypi.org/p/pyaccesskit) from GitHub Actions with **Trusted
Publishing**: PyPI trusts this repository's `release.yml` workflow directly, so no API token is stored
anywhere. The documentation is deployed to GitHub Pages by `docs.yml`.

!!! warning "Uploads are permanent"
    A version number can be uploaded to PyPI only once. It cannot be replaced, even after deleting it.
    Rehearse on TestPyPI whenever something about packaging changed.

## One-time setup

1. **PyPI and TestPyPI accounts.** Create them on [pypi.org](https://pypi.org/account/register/) and
   [test.pypi.org](https://test.pypi.org/account/register/). They are separate, and both need two-factor
   authentication.
2. **Pending trusted publishers.** This registers the project before its first upload. On each site, go to
   *Your account → Publishing → Add a new pending publisher* and enter:

    | Field | Value |
    |---|---|
    | PyPI project name | `pyaccesskit` |
    | Owner | `ariedotcodotnz` |
    | Repository name | `PyAccessKit` |
    | Workflow name | `release.yml` |
    | Environment name | `pypi` (on TestPyPI: `testpypi`) |

3. **GitHub environments.** In the repository, go to *Settings → Environments* and create `pypi` and
   `testpypi`. On `pypi`, add yourself as a required reviewer, so each upload waits for your approval.
4. **GitHub Pages.** Go to *Settings → Pages → Source: GitHub Actions*. The docs deploy on every push to
   `main`.

## Each release

1. **Update the version** in `src/pyaccesskit/_version.py`, following
   [semantic versioning](https://semver.org/). Until 1.0, a minor release may break the API; say so in
   the changelog.
2. **Update `CHANGELOG.md`.** Move the entries from *Unreleased* to a new `## [X.Y.Z] - YYYY-MM-DD` section,
   and update the comparison links at the bottom.
3. **Run the full checks on a Windows machine with Access.** CI cannot run the real-Access tests:

    ```console
    uv run ruff check . && uv run ruff format --check . && uv run pyright
    uv run python scripts/sync_docs.py --check
    uv run pytest --integration
    UV_PROJECT_ENVIRONMENT=.venv-x86 uv run --python cpython-3.12-windows-x86 pytest --integration
    ```

4. **Rehearse (optional).** Go to *Actions → Release → Run workflow → target: testpypi*. Then install from
   TestPyPI in a clean environment:

    ```console
    uv venv /tmp/pak && uv pip install --python /tmp/pak --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple/ pyaccesskit
    ```

5. **Commit, push, and publish a GitHub release** with the tag `vX.Y.Z`. The tag must match the package
   version, or the workflow stops. The workflow runs the checks, builds the sdist and wheel, and uploads
   them to PyPI after you approve the `pypi` environment.
6. **Check** [pypi.org/p/pyaccesskit](https://pypi.org/p/pyaccesskit), then `pip install -U pyaccesskit`
   and `pyaccesskit doctor --probe` on a machine with Access.

## Publishing without GitHub Actions

With an API token from *PyPI → Account settings → API tokens*:

```console
uv build
uvx twine check --strict dist/*
uv publish --publish-url https://test.pypi.org/legacy/ --token pypi-…   # rehearsal on TestPyPI
uv publish --token pypi-…                                                # PyPI
```

Prefer a token scoped to the `pyaccesskit` project, and keep it out of shell history. Set it in the
`UV_PUBLISH_TOKEN` environment variable instead of passing `--token`.

# Changelog

All notable changes to the Glassnode Excel add-in. The version is the one the task pane
shows (`src/version.ts`). Each release is tagged `v<version>` on `main`, and the CDN
serves the build of that tag. Security fixes are listed under **Security**.

## 0.3.0 (unreleased: set the date in the release PR just before it merges)

### Changed

- API calls go through the `glassnode-api` TypeScript client instead of axios. Failed
  requests (429, 5xx, network errors) are retried twice, and at most 6 requests run at once.
- Responses are cached in memory for 3 minutes (metrics) and 1 hour (assets); a failed
  request is never cached. The asset list is checked against the API's full asset shape, so
  a malformed list shows "Error: Invalid response format".
- A parameter such as `f=xml` now shows an input error instead of being sent.

### Security

- Runtime dependencies cut from axios, axios-cache-interceptor and their dependencies to the
  `glassnode-api` client (Glassnode's own, published with npm provenance) and zod. The client
  masks the key in its errors and doesn't follow redirects, so the key isn't sent to another
  host.

## 0.2.0 (2026-11-27)

### Security

- The API key is sent in the `X-Api-Key` header instead of the `api_key` query parameter,
  so it no longer appears in request URLs, proxy logs or browser history.
- The settings field masks the key, and the saved key is never shown again (only its last
  four characters). A new **Remove API Key** button deletes it from the add-in's storage.
- Failed requests log only the error code, status and message, never the request with
  the key.
- A formula can no longer set `api_key` as a parameter, and metric paths may only contain
  letters, digits, `_` and `-`. A shared workbook can't send the key to other API
  endpoints.
- axios is upgraded from 1.11.0 to 1.21.1 and axios-cache-interceptor from 1.8.0 to
  1.13.0, fixing known vulnerabilities (high severity in axios and its form-data
  dependency). core-js is updated to 3.51.0.

- Formula calls are logged without parameter values, so a mistyped `api_key=...` never
  reaches the console.
- `pick` and object columns read only the object's own keys, never inherited ones such as
  `constructor`.

### Changed

- `METRIC` rejects `a`, `s`, `u`, `source` and `metric` as `key=value` parameters, since
  its own arguments set them. `i` (resolution) can still be set. A parameter value may now
  contain `=`.
- Dates that don't exist (such as `2024-13-01`, which used to be read as a 1905 date) and
  serial numbers outside Excel's range return the invalid-date error. An end date of
  1970-01-01 is no longer ignored.
- `ASSETS` returns an error unless `limit` is a whole number of 0 or more.
- Objects and arrays from the API are written to cells as text, and empty values as empty
  cells.
- Cached responses are cleared when the API key changes, so data fetched with one key is
  never shown under another.
- Build and development dependencies updated.
- The build writes `version.json` with the version, commit and tag, so the CDN shows which
  build it serves.

### Development

- Install with `npm ci --ignore-scripts`. Node.js 26.13.1 (`.nvmrc`). The unused
  `request` package is removed.
- Dependabot proposes weekly dependency updates. Pull requests use a template with a
  security-review line.

## 0.1.2 and earlier

Not recorded.

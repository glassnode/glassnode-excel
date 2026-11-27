# Changelog

All notable changes to the Glassnode Excel add-in. The version is the one the task pane
shows (`src/version.ts`). Each release is tagged `v<version>` on `main`, and the CDN
serves the build of that tag. Security fixes are listed under **Security**.

## 0.2.0 (not yet released: set the date when tagging)

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
  1.11.1, fixing known vulnerabilities (high severity in axios and its form-data
  dependency).

### Changed

- Cached responses are cleared when the API key changes, so data fetched with one key is
  never shown under another.
- Build and development dependencies updated.

### Development

- Install with `npm ci --ignore-scripts`. Node.js 26.13.1 (`.nvmrc`). The unused
  `request` package is removed.
- Dependabot proposes weekly dependency updates. Pull requests use a template with a
  security-review line.

## 0.1.2 and earlier

Not recorded.

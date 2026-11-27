# `glassnode-excel`: agent instructions

Read [DEVELOPMENT.md](DEVELOPMENT.md) first: setup, commands and the release steps live
there. Before opening a PR, run `npm ci --ignore-scripts`, `npm run build`, `npx jest` and
`npm run validate`, and fill in the PR template, including its security-review line. Never
approve or merge your own PR: a code owner approves and merges it.

## Releases

- Every release bumps `ADDIN_VERSION` in `src/version.ts` and adds its entry to
  [CHANGELOG.md](CHANGELOG.md), with security fixes under **Security**. A PR that changes
  what users get adds its line to the changelog's next version.
- No tag, no deploy: the CDN serves only a build of a `v<version>` tag on `main`, built
  from a clean checkout of that tag (DEVELOPMENT.md → Releasing).

## Don't

- Run `deploy.sh` or write to the CDN bucket: a human deploys.
- Change `<Version>`, `<AppDomains>`, the hosts or the permissions in `manifest.xml`
  without a discussion first.
- Put the API key anywhere but the `X-Api-Key` header: not in a URL, a log line, an error
  message or the task pane.
- Add a runtime dependency (it ships to every user) without a discussion first.

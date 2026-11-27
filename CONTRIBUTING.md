# Contributing

## Contents

- [Build and test](#build-and-test)
- [Commit messages](#commit-messages)
- [Making a change](#making-a-change)
- [Releasing](#releasing)

Running the add-in locally in Excel (dev server, sideloading, debugging) is in
[DEVELOPMENT.md](DEVELOPMENT.md).

## Build and test

Node is the version in [`.nvmrc`](.nvmrc).

```bash
npm ci --ignore-scripts
npm run lint:check
npx prettier --check 'src/**/*.{ts,tsx}'
npm test
npm run build
npm run validate
```

CI runs the same checks, except `validate`, on every PR. `--ignore-scripts` stops packages
from running install scripts on your machine; pass it to `npm install` too.

## Commit messages

Follow [Conventional Commits](https://www.conventionalcommits.org/en/v1.0.0/):
`<type>(<optional scope>): <summary>`, e.g. `fix(metric): reject impossible dates`. Types:
`feat`, `fix`, `perf`, `refactor`, `docs`, `test`, `ci`, `chore`. A breaking change adds
`!` (`feat!: …`) or a `BREAKING CHANGE:` footer. The version bump: breaking major, `feat`
minor, any other change to what users get patch. The PR title follows the same format, and
the squash commit takes the PR title.

## Making a change

1. Open a PR against `main` and fill in the PR template. CI must pass.
2. Bump `ADDIN_VERSION` in [`src/version.ts`](src/version.ts) following semantic
   versioning. The task pane shows it. (`version` in `package.json` isn't used.)
3. Add an entry for that version at the top of [CHANGELOG.md](CHANGELOG.md), with
   security fixes under **Security**.
4. A code owner approves and merges.

Steps 2 and 3 apply to changes to what ships in the add-in; skip them for docs-, CI- and
tooling-only changes, and for changes only to development dependencies or the lockfile.
Several PRs can go into one release: each adds its line to the changelog's next version,
and the release PR sets the version and the date. On a Dependabot runtime PR, push steps 2
and 3 onto it before merging.

## Releasing

No tag, no deploy: the CDN serves only a build of a `v<version>` tag on `main`.

1. Merge the release PR (version and dated changelog heading).
2. A code owner with the Maintain or Admin role tags the squash commit on `main` and pushes
   the tag. `v*` tags can't be moved or deleted afterwards, so a fix ships as a new version.
   ```bash
   git switch main && git pull
   git tag -a v<version> -m "v<version>" && git push origin v<version>
   ```
3. Build from a clean checkout of the tag, never from a working copy:
   ```bash
   git switch --detach v<version>
   npm ci --ignore-scripts
   npm run build && npm run validate
   cat dist/version.json   # the tag, and "dirty": false
   ```
4. Test that build in Excel (web and Windows) by sideloading it.
5. A Glassnode maintainer deploys `dist` to the CDN with `deploy.sh`. The CDN caches for up
   to an hour. Afterwards `https://cdn-msoffice.glassnode.com/msexcel/version.json` shows
   the tag.
6. Create the GitHub Release for the tag, with the `CHANGELOG.md` section as its notes.
7. Only when `manifest.xml` changed: submit the new manifest in Microsoft Partner Center.
   Microsoft reviews it before users get it.

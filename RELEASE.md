# Releasing Airy

This document describes how a version of Airy is cut, what ends up in a
GitHub release, and how installed apps pick updates up. The release pipeline
is the `Release` workflow (`.github/workflows/release.yml`); it packages
Linux, Windows, and macOS installers and publishes them through
electron-builder's GitHub provider (`publish: { owner: besliky, repo: airy }`
baked into `apps/shell/electron-builder.cjs`).

## Cutting a release

The app version lives in `apps/shell/package.json` (`version`). A release is:

```bash
# 1. bump the version (nothing else tracks it)
npm version patch -w @airy-office/shell   # or minor / major

# 2. commit the bump, then tag it with the SAME version
git tag v0.9.1

# 3. push the commit and the tag
git push origin main v0.9.1
```

Pushing the `v*` tag triggers the Release workflow on `ubuntu-latest`,
`windows-latest`, and `macos-latest`. Each job regenerates the third-party
notices, builds all six apps, and runs electron-builder with
`--publish always`, which uploads the artifacts to the release named after
the tag. Watch the run under the repository's **Actions** tab.

electron-builder creates that release as a **draft**. Drafts are invisible
to the in-app updater (see below), so the last manual step is: open
**Releases**, review the draft's assets, and publish it.

### The tag must match `apps/shell`'s version

electron-builder publishes under `v{version}` taken from `apps/shell/package.json`,
not from the pushed tag. If the tag and the package version disagree, the run
uploads to a release for `v{package-version}` while you pushed a different
tag — leaving a stray draft and an empty-looking release. Always tag exactly
what `apps/shell` says.

### Manual runs (no publishing)

`workflow_dispatch` runs the same two jobs but replaces publishing with
`--publish never`; the installers are attached to the workflow run as
artifacts (`airy-linux`, `airy-windows`) and nothing touches GitHub
Releases. Use this to smoke-test packaging on a branch. The run's `publish`
input opts a manual run into publishing (a draft release is created for the
current `apps/shell` version) — handy for a release candidate you intend to
hand out but not announce.

## What lands in a release

| Artifact                                    | Platform | Purpose                                              |
| ------------------------------------------- | -------- | ---------------------------------------------------- |
| `Airy Setup <v>.exe`                        | Windows  | NSIS installer (assisted, install dir selectable)    |
| `latest.yml` + `.exe.blockmap`              | Windows  | update feed + differential-download metadata         |
| `Airy-<v>.AppImage`                         | Linux    | self-contained x64 AppImage                          |
| `latest-linux.yml` (+ `.AppImage.blockmap`) | Linux    | update feed + blockmap for the AppImage              |
| `airy_<v>_amd64.deb`                        | Linux    | apt package (`packageName: airy`, upgrades in place) |
| `Airy-<v>-arm64.dmg`                        | macOS    | arm64 disk image (x64 builds are opt-in, see below)  |
| `Airy-<v>-arm64-mac.zip`                    | macOS    | zip variant of the same arm64 build                  |

The `latest*.yml` files are what electron-updater reads; they are release
assets just like the binaries and must not be deleted from a published
release, or in-app updates stop resolving.

## How updates reach installed apps

The updater (`apps/shell/src/main/updater/`) checks the feed ~10 s after
launch and via **Help → Check for Updates**. Behavior per install type:

- **Windows (NSIS)** — in-app: dialog offers the new version, downloads with
  progress, then restarts into the installer (or applies on quit).
- **Linux (AppImage)** — in-app, same flow; the new AppImage replaces the
  running one.
- **Linux (deb)** — notify-only: electron-updater cannot self-update a deb,
  so the app raises a notification linking to the
  [latest release](https://github.com/besliky/airy/releases/latest); the user
  downloads `airy_<v>_amd64.deb` and installs it (`apt install ./…` upgrades
  the `airy` package).

macOS has no in-app update path (unsigned builds make auto-update
trust-reduced, so no feed check runs there): Help → Check for Updates shows
a dialog linking to the releases page, and a new version arrives by
downloading the new `Airy-<v>-arm64.dmg` from there.

## Manually verifying the update flow

Before announcing a release, walk one cycle end to end:

1. Install version X (the NSIS exe on Windows, the AppImage or deb on Linux)
   from a published release — drafts do not count, the updater cannot see
   them.
2. Cut and publish version Y (X+1) as described above.
3. Relaunch the X install:
   - **NSIS / AppImage** — within ~10 s (or immediately after Help → Check
     for Updates) a dialog offers `Y` with a Release Notes button. Choose
     Download, watch progress reach 100 %, then install now: the app quits,
     updates, relaunches on Y, and a repeat check reports up-to-date.
   - **deb** — a notification appears linking to the releases page; the app
     itself stays on X until the new deb is installed.
4. A failed check (airplane mode, rate limit) must stay quiet: no dialog,
   just a transient status in the Help menu.

## Notes

- The fork ships unsigned installers. Windows shows a SmartScreen/"unknown
  publisher" prompt on first run and macOS Gatekeeper warns about the
  unsigned dmg; that is expected for now.
- macOS releases build arm64 only by default; an Intel (x64) variant is
  opt-in via `AIRY_MAC_X64=1` in the workflow environment (see the mac
  block in `apps/shell/electron-builder.cjs`).
- The rpm target is intentionally not built by the release workflow; the
  local `npm run dist:linux` still packages it when `rpmbuild` is present.

# Contributing

Use Node.js 24 and `npm ci`. Edit `src/`, run `npm run check`, and commit both source and generated `theme.css`. Keep native fonts, keyboard focus, Source mode and clean Markdown intact. Do not add network assets, `!important`, `:has()`, runtime scripts or disabled lint rules.

Text files use LF. Markdown examples intentionally use two trailing spaces for hard line breaks; preserve those spaces.

Changes go through a pull request and CI before a version tag. Keep package, manifest, versions and changelog synchronized. A tag named exactly like the manifest version triggers the release workflow and uploads `theme.css`, `manifest.json` and the example ZIP.

## Optional native QA

Only the canonical development vault `@Nexus-dev` is used. Set `OBSIDIAN_DEV_ROOT` to the workspace containing that vault when this checkout is elsewhere; the default is two directories above the repository. `OBSIDIAN_CDP` defaults to `http://127.0.0.1:9235` and accepts localhost only.

Install the built theme in that vault and open Obsidian with local debugging enabled. Run `npm run qa:obsidian`. The scripts verify the vault path, write only known QA fixtures under `Pruebas/`, restore edited fixture content and font/color preferences, and store results in ignored `.qa/`. Existing examples are read and hashed, never overwritten. Do not use a personal or production vault.

Native QA is separate from CI: a passing CSS build is not a visual test, and viewport emulation does not establish physical-device support.

`npm run qa:layout` checks the four views at 320, 390 and 1400 pixels, in both modes and both color schemes, and refreshes the native screenshots in `images/`.

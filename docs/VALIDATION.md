# Office Mock 1.0.0 validation

Validated on 2026-11-01 in Windows desktop Obsidian **1.14.5**, using the canonical `@Nexus-dev` vault and a separate application profile. Community plugins were not enabled in that profile. Office Mock was selected as the actual theme; both former Office snippets were disabled.

## Automated checks

- Official `stylelint-config-obsidianmd`: zero errors and zero warnings, without disabling rules.
- Additional browser compatibility lint: Chromium 114, zero warnings.
- Build and static checks: 267 CSS rules, 107 offline SVG uses, required manifest fields, matching versions and complete project/third-party licenses.
- No remote imports/resources, runtime JavaScript, imposed font families or text size, `!important` or `:has()` in the distributed theme.
- Twelve foreground/background token pairs meet at least 4.5:1 contrast. The 512 × 288 PNG thumbnail is checked directly.
- Four examples and four templates have normal frontmatter, headings and Markdown content, fictional `.example` email addresses and no UI HTML, special UI callouts or decorative empty rows.
- The installation ZIP is decompressed and every entry is compared byte-for-byte with its source. It includes the theme, examples, templates, documentation and screenshots, with no scripts or vault configuration.

## Native integration: 100 checks

The reading suite passed **77 checks**. It covers all eight notes in light/dark modes, native rendering of offline SVG assets, property display settings, default document layout, the `office-plain` opt-out, transclusions without nested chrome, narrow panes, horizontally scrolling data, Source mode and switching themes. All example and fixture hashes were preserved.

The editing suite passed **23 checks**. It covers all four views in light/dark Live Preview, configured text and interface fonts, text input and saving to Markdown, native undo, activating a real cell editor, typing two cell values, `Tab` and `Shift+Tab`, persisted data on disk, reading back edited values and horizontal table scrolling. Edited QA fixtures, temporary fonts and color preferences were restored.

## Layout and visual review: 48 checks

All four views were checked in reading and Live Preview, in light/dark modes, with viewports of **320, 390 and 1400 pixels**. Checks cover pane overflow, the colored header, document foreground/surface tokens and the actual rendered heading contrast. These are viewport checks, not physical-device certification.

Visual review found and corrected a fixed-color document subheading and a white SVG page background that showed hairlines between dark sections. The final layout checks and screenshots were regenerated after those corrections.

Sixteen screenshots in `images/` show reading and editing for each application in both modes. The public screenshots crop the native note view so other vault notes are not included. The thumbnail is captured from the actual document view, not a design mockup.

## Scope

Native QA does not claim testing on every historical Obsidian version, operating system, community plugin or physical mobile device. Minimum app version 1.5.0 corresponds to native table editing; the installer engine must support the CSS features checked against Chromium 114.

Decorative ribbons/search/formula bars remain static. Formula calculation, email delivery, scheduling and native Office file formats are outside the theme. Native text/table editing remains available and the original Markdown is preserved.

A GitHub CI/release check and acceptance by Community Themes are separate. The theme is prepared for submission; no community directory acceptance is claimed.

## 1.0.1 license metadata correction

The project license now matches GitHub's standard MIT template, checked against the fixture in `tests/mit-template.txt`. Third-party notices are preserved byte-for-byte and remain embedded in the distributed CSS. The CSS rules excluding comments were compared with release 1.0.0 and are identical. Static checks, official lint, browser compatibility and all ZIP entry comparisons pass for 1.0.1; the native UI validation above remains the evidence for the unchanged styling.

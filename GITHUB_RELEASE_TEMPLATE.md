# GitHub Open Source Release Template

The text below can be used as a starting point for the repository homepage,
the first GitHub release note, or a public announcement.

---

## Repository Description Template

yuninput is a Chinese input method project for Windows 11 TSF (Text Services Framework).

The repository currently provides:

- A C++ TSF Text Service framework
- An input engine capable of loading external code tables
- Candidate window support and basic user learning features
- Configuration tooling, installation scripts, and EXE/MSI packaging flows

This project is under active development. Its goal is to provide an extensible,
maintainable, and verifiable Chinese input method framework, with progressive
support for Zhengma-style input experiences.

## Open Source Statement Template

The original code in this repository is released by Intel Corporation.

Unless otherwise stated, the project code is licensed under GPL-3.0-or-later.
See [LICENSE](LICENSE).

For ownership, third-party notices, Fcitx-related materials, and redistribution
requirements, see:

- [NOTICE.md](NOTICE.md)
- [NOTICE.zh-CN.md](NOTICE.zh-CN.md)
- [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)
- [THIRD_PARTY_NOTICES.zh-CN.md](THIRD_PARTY_NOTICES.zh-CN.md)

## Zhengma-related Statement Template

This project may contain implementations, import workflows, conversion logic,
or documentation related to Zhengma-compatible input experiences.
This project does not claim ownership of the Zhengma input method system,
related names, standards, trademarks, or other intellectual property.
Such rights belong to their respective rights holders.

## Third-Party Dictionary Statement Template

This project supports importing external open-source code tables and currently
keeps some Fcitx-format source table snapshots for conversion and validation.
The copyright and license of those third-party materials remain with their
upstream authors and projects.

Before any public redistribution, the maintainer should verify and preserve the
exact upstream links, license texts, and attribution requirements.

## Current Limitations Template

- This project is still under development and should not be treated as a final commercial product.
- Redistribution rights for some dictionaries or derived dictionaries depend on upstream license status.
- If some third-party dictionaries cannot be redistributed, users must import them separately.

## First Release Note Template

This release includes:

- The Windows TSF input method core DLL
- The configuration tool
- Install and uninstall scripts
- An EXE installer
- A wrapped MSI installer

Known notes:

- The project is still evolving, and behavior, interfaces, and dictionary organization may change.
- Distribution boundaries for third-party dictionaries are governed by the notice files in this repository.

Recommended final checks before publication:

1. Verify the exact upstream source and license of each file under `third_party/fcitx/`.
2. Remove any non-public diagnostics, logs, or machine-specific paths.
3. Run one final build, install verification, and artifact hash archival.

# Third-Party Notices

This repository includes code, scripts, and data processing logic authored for yuninput.
Some dictionary inputs may be generated from external open-source projects or public resources.

## Intel Ownership Statement

- Unless otherwise noted, the original code in this repository is copyrighted by Intel Corporation.
- Third-party materials remain under the ownership of their respective rights holders.

## Fcitx-related Sources

- Some dictionary conversion workflows in this repository can import source tables from Fcitx-related projects.
- The original copyright and license of those source tables belong to their original authors and projects.
- This repository does not relicense third-party source tables.

The repository currently contains vendored Fcitx-format table snapshots under:

- `third_party/fcitx/zhengma.txt`
- `third_party/fcitx/zhengma-large.txt`
- `third_party/fcitx/zhengma-pinyin.txt`

These files are identifiable as Fcitx table files by their file headers such as:

- `;fcitx Version 0x03 Table file`
- `;fcitx 版本 0x03 码表文件`

At the time of writing, this repository records the source format and source family, but does not embed a separate upstream license text alongside those vendored snapshots.
Before a public GitHub release, the maintainer should verify and record:

1. The exact upstream repository or distribution package from which each file was obtained.
2. The applicable upstream license identifier and license text.
3. Any attribution or redistribution conditions attached to that upstream source.

If you distribute dictionary files derived from Fcitx-related sources, you must:

1. Keep the original attribution and license notices from the source project.
2. Comply with the original source license terms.
3. Provide source links and license texts in your distribution package.
4. If the upstream license cannot be verified, do not redistribute the vendored source table.

## Zhengma Rights Statement

- The rights of the Zhengma input method system, related standards, names, and associated intellectual property belong to their respective rights holders.
- This project only implements an open-source input method framework and does not claim ownership of Zhengma intellectual property.
- If any name, trademark, or material should be adjusted, please open an issue and the project will cooperate with correction.

## Distribution Checklist

Before publishing releases to GitHub:

1. Verify every imported dictionary source has a clear, compatible license.
2. Keep upstream copyright and license information.
3. Remove files with unknown or non-redistributable license status.
4. Update this file with concrete upstream links and license identifiers.

For a Chinese version of this notice, see THIRD_PARTY_NOTICES.zh-CN.md.

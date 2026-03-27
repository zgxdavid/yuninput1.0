# Open Source Release Checklist

This checklist is intended for the final review before publishing yuninput to GitHub.

## 1. Ownership and Approval

1. Confirm that the original code in this repository may be released by Intel.
2. Confirm that all required internal review, legal review, and approval steps are complete.
3. Confirm that no internal-only documents, credentials, email addresses, private paths, or debug artifacts remain in the repository.

## 2. License Files

1. Confirm that the repository root contains `LICENSE`.
2. Confirm that the copyright holder and license choice match the intended release policy.
3. Confirm that `NOTICE.md` and `NOTICE.zh-CN.md` accurately state Intel ownership.

## 3. Third-Party Materials

1. Review `third_party/`, `data/`, `tools/`, and `scripts/` for third-party materials.
2. Record, for each third-party item, its source URL, upstream project name, license, and redistribution status.
3. Remove or replace any material with unknown license status or non-redistributable terms.
4. Update `THIRD_PARTY_NOTICES.md` and `THIRD_PARTY_NOTICES.zh-CN.md`.

## 4. Fcitx-related Verification

1. Verify the upstream source and license for `third_party/fcitx/zhengma.txt`.
2. Verify the upstream source and license for `third_party/fcitx/zhengma-large.txt`.
3. Verify the upstream source and license for `third_party/fcitx/zhengma-pinyin.txt`.
4. If the upstream license of any of these files cannot be verified, do not redistribute that file or its direct derivatives.
5. If these files remain in the public repository, add exact upstream links and license texts to the notice files.

## 5. Zhengma-related Statements

1. Confirm that Zhengma rights statements remain present in `NOTICE.md`, `NOTICE.zh-CN.md`, and `THIRD_PARTY_NOTICES.md`.
2. Confirm that no public wording implies ownership of Zhengma intellectual property.
3. If the release uses the term "Zhengma", ensure it is used only as a factual compatibility or descriptive reference.

## 6. Repository Cleanup

1. Check for temporary logs, personal paths, debug snapshots, install logs, or machine-specific generated files.
2. Review whether content under `diagnostics/` is appropriate for public release; remove or relocate it if needed.
3. Review install, registration, and packaging scripts for internal-only paths or assumptions.
4. Review README, docs, and comments for non-public information.

## 7. Release Artifacts

1. Confirm that the source release, EXE, and MSI match the repository notices.
2. Confirm that installer packages include the required license and notice files.
3. Confirm that no installer package contains dictionaries, assets, or resources with unclear redistribution rights.

## 8. GitHub Release Page

1. Describe the project scope, maturity, and intended usage on the repository homepage.
2. Include license information, third-party notice summary, and known limitations in the release notes.
3. If some dictionaries cannot be redistributed, state that users must import them separately.

## 9. Final Gate

1. Run one final full build and installation verification.
2. Regenerate release artifacts and record hashes.
3. Complete this checklist before making the repository public.

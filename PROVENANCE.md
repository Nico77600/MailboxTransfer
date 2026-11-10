# Public provenance declaration

- Repository: https://github.com/Nico77600/MailboxTransfer
- Default branch: `main`
- Release: `v1.1.0` (`MailboxTransfer-1.1.0.zip`)
- Base of the release branch: public `main` at `95fd6822ca56becbab9923573207505b97c106a4` (tag `v1.0.1`, merge of pull request #2, 2026-11-10)
- Audited revision: `da2dc82afef08a3eb9296e5d9a061115f5787e8e` (public `main` and the tag `v1.1.0`, merge of pull request #3; this provenance record is the next commit)
- Audit date: 2026-11-10
- Public commits inspected across branches and tags: 16 (the branches `main`, `publish/mailboxtransfer-1.0.0`, `release/1.0.1-20261110` and `release/1.1.0-20261110`, and the tags `v1.0.0`, `v1.0.1` and `v1.1.0`)
- `Co-authored-by: Copilot` findings: none (none of the 16 commits carries a `Co-authored-by` trailer at all)
- Previous audit: `aa916558bc0368e5852a0ec4a9306044942ea150`, 2026-11-10, 9 public commits, no `Co-authored-by: Copilot` finding

This declaration records the public commit-history audit at the revision above. It is not a cryptographic signature and does not make claims about how source code was authored. Release assets are covered by the GitHub Actions workflow in `.github/workflows/attest-release.yml`; run it manually for an existing release when needed.

## Package integrity

The release ZIP (`MailboxTransfer-1.1.0.zip`, built by `tools/New-MailboxTransferPackage.ps1`) is the user deliverable; the source archives of a release are not.

- Every PowerShell file of the package, `src/Test-PackageIntegrity.ps1` and the manifest `PackageIntegrity.psd1` carry the owner's Authenticode signature with a trusted timestamp. The verifier hard-codes the signer thumbprint (`0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6`).
- The manifest lists the SHA-256 of every other protected file: the C# sources compiled in memory, the SQLite libraries, the window definitions, the HTML guides, the launcher, the README and the notices. The launcher (`.cmd`) cannot carry an Authenticode signature: the manifest protects it by hash.
- The entry scripts and the module run the verifier before anything is compiled or shown; a file that is missing, changed, unknown or a reparse point stops the start. The configuration and the generated `out\` and `logs\` folders are the only unprotected places.
- Continuous integration verifies the signed sources with `-DevelopmentMode` (files and signatures, no unsigned mode in a release). No key material is read by any tool of the repository.

## Attesting the ZIP

`.github/workflows/attest-release.yml` downloads the published ZIP assets of an existing release tag and attests the ZIP bytes. Authenticode signatures are validated on Windows, at launch, by the verifier.

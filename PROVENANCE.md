# Public provenance declaration

- Repository: https://github.com/Nico77600/MailboxTransfer
- Default branch: `main`
- Release: `v1.0.0` (`MailboxTransfer-1.0.0.zip`)
- Audited revision: `05b75b5cef113fa54ba17fc7c620f9ca32647432` (public `main` and the tag `v1.0.0`, merge of pull request #1; this provenance record is the next commit)
- Audit date: 2026-11-10
- Public commits inspected across branches and tags: 4 (the branches `main` and `publish/mailboxtransfer-1.0.0`, and the tag `v1.0.0`)
- `Co-authored-by: Copilot` findings: none (none of the 4 public commits carries a `Co-authored-by` trailer at all)

This declaration records the public commit-history audit at the revision above. It is not a cryptographic signature and does not make claims about how source code was authored. Release assets are covered by the GitHub Actions workflow in `.github/workflows/attest-release.yml`; run it manually for an existing release when needed.

## Package integrity

The release ZIP (`MailboxTransfer-1.0.0.zip`, built by `tools/New-MailboxTransferPackage.ps1`) is the user deliverable; the source archives of a release are not.

- Every PowerShell file of the package, `src/Test-PackageIntegrity.ps1` and the manifest `PackageIntegrity.psd1` carry the owner's Authenticode signature with a trusted timestamp. The verifier hard-codes the signer thumbprint (`0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6`).
- The manifest lists the SHA-256 of every other protected file: the C# sources compiled in memory, the SQLite libraries, the window definitions, the HTML guides, the launcher, the README and the notices. The launcher (`.cmd`) cannot carry an Authenticode signature: the manifest protects it by hash.
- The entry scripts and the module run the verifier before anything is compiled or shown; a file that is missing, changed, unknown or a reparse point stops the start. The configuration and the generated `out\` and `logs\` folders are the only unprotected places.
- Continuous integration verifies the signed sources with `-DevelopmentMode` (files and signatures, no unsigned mode in a release). No key material is read by any tool of the repository.

## Attesting the ZIP

`.github/workflows/attest-release.yml` downloads the published ZIP assets of an existing release tag and attests the ZIP bytes. Authenticode signatures are validated on Windows, at launch, by the verifier.

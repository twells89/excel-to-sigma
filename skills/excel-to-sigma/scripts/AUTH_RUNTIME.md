# Vendored Sigma auth runtime

The files below are copied **byte-for-byte** from
`github.com/twells89/sigma-skills`, commit `95e67b7`, directory `sigma-api/`
(`scripts/*`, `scripts/lib/*`, `reference/browser-oauth-login.md`). They were also
compared with the `sigma-migration-skills` shared runtime (`shared/scripts/`,
`get_token.py` last changed in `334752a6`): `get_token.py`, `get-token.sh`,
`browser-login.sh`, `refresh-token.sh` and `lib/browser-login-platform.sh` are
identical to the sigma-skills copies, so there is a single canonical runtime.

Do not edit them here. To update, re-vendor from the source repo and change the
hashes below in the same commit; `tests/test_sigma_auth.py` fails on any drift so a
change is always deliberate.

`scripts/lib/sigma_auth.py` is this skill's own resolver (not vendored); it calls
`get_token.mint_token()` as its last resolution step.

| File | SHA-256 |
|---|---|
| scripts/get_token.py | 1967ece24a8eae8489ac2ee922ec9d26e79e9d5e7b9e63445c1a785768d0d658 |
| scripts/get-token.sh | 9d951f86806d4f835cc8ba45766d7f715f5062bf31981816c252b328122bc76b |
| scripts/browser-login.sh | 64f3d5248e32a59f89765f9e106e238574cd45dbc2460dc6f84a676aab9f8a7e |
| scripts/refresh-token.sh | acce747146e6c8eadff609e062811801c0f75eb5211003a6d3ae9f88a353fe46 |
| scripts/lib/browser-login-platform.sh | dae84d7c1d1af9433f4ae81882952c0e0b8240f467b5620ad9af9115e85303e7 |
| refs/browser-oauth-login.md | ecba73c674bf1abd4bcfba63b74088a7075174dd8e60bd8e3389721367112637 |

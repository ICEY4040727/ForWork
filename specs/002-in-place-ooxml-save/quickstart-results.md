# Quickstart Results: In-Place OOXML Save

**Feature**: 002-in-place-ooxml-save

**Recorded**: 2026-11-18

**Environment**: Windows 11, Node v24.15.0, pnpm 9.15.9

**Scope**: T037, T046, T049, T050, T058

## T049. SDK decision

**Decision**: no checked-in TypeScript or Python scenario emits `document/save`, so this change adds no scenario and changes no expected output. If a later scenario adds document events, update both SDK projections in the same change.

Both expected trees exist and were verified:

| Tree | Tracked files |
| --- | --- |
| `packages/sdk/client/tests/expected/` | `cot-translation-projection.json` |
| `python/sdk/tests/expected/` | `account-provider-signout.json`, `cot-translation-projection.json`, `team-mode-projection.json` |

Commands and outcomes:

```sh
rg -n "document/save" packages/sdk python/sdk --glob "!**/node_modules/**" --glob "!**/lib/**" --glob "!**/.venv/**"
```

Outcome: exit 1, no matches. No TypeScript or Python SDK file emits `document/save`.

```sh
rg -n "document/save" packages/sdk/client/src packages/sdk/client/tests python/sdk/src python/sdk/tests
```

Outcome: exit 1, no matches. The scenario and expected trees contain no `document/save` event.

```sh
git status --porcelain -- packages/sdk/client/tests/expected python/sdk/tests/expected
```

Outcome: no output; 0 changed, staged, or untracked paths.

```sh
git diff --exit-code -- packages/sdk/client/tests/expected python/sdk/tests/expected
```

Outcome: exit 0, no diff. Both expected trees are unchanged.

```sh
git ls-files packages/sdk/client/tests/expected python/sdk/tests/expected
```

Outcome: the four tracked files listed in the table above.

## T037. Macro-enabled provider saves

`packages/document/office-exchange/tests/provider-docm.spec.ts` calls `readDocumentWithSource`/`writeDocumentFromSource` and `readWorkbookWithSource`/`writeWorkbookFromSource` directly, so the provider splice path runs without the Remote whole-file save gate.

For `macro.docm` and `macro.xlsm`, the test edits a document heading and a worksheet cell respectively, writes from the source package, and asserts:

- the `vbaProject.bin` entry is byte-identical;
- the macro-enabled main content type (`application/vnd.ms-word.document.macroEnabled.main+xml` or `application/vnd.ms-excel.sheet.macroEnabled.main+xml`) is present and unchanged;
- the output contains no macro-free main content type;
- the extension/content-type pair is not swapped: a `.docm` output keeps the Word main part and Word content types, and a `.xlsm` output keeps the workbook main part and Excel content types;
- the VBA relationship and VBA content-type override are intact;
- the edited part changed and the disposition does not name the VBA entry.

```sh
pnpm exec vitest run packages/document/office-exchange/tests/provider-docm.spec.ts
```

Outcome: 1 test file passed, 2 tests passed.

```sh
pnpm exec vitest run packages/document/office-exchange/tests/ooxml-preservation.spec.ts -t "keeps the VBA project"
```

Outcome: 1 test file passed, 2 tests passed, 89 skipped.

## T050. Projection audit

The audit found the new optional `disposition` remains audit-only:

- `documentProjectionEntrySchema` has no `disposition` field;
- `foldDocumentProjection` reads only `documentId` and `contentVersion` from `document/save`;
- `sameEntry` compares the existing entry fields only;
- `stateVersion` is already 2 for the macro-enabled format union;
- `SESSION_FORMAT_VERSION` is unchanged.

`packages/session/session-document/tests/document-projection.spec.ts` asserts a repeated save at the same version with a disposition returns the same state reference, that the entry has no disposition field, that the entry round-trips through the state schema, and that `stateVersion` is 2.

Required verification:

```sh
pnpm exec vitest run packages/document/office-exchange/tests/provider-docm.spec.ts packages/session/session-document/tests
```

Outcome: 8 test files passed, 55 tests passed.

```sh
pnpm exec tsc -b packages/document/office-exchange packages/session/session-document
```

Outcome: exit 0, no diagnostics.

## T058. Large-workbook save performance

`packages/document/office-exchange/tests/ooxml-preservation.spec.ts` builds the workload in memory with ExcelJS: 10,000 numeric cells in one worksheet, plus a 4,500,000-byte pseudo-random `xl/media/perf.bin` entry declared through a `bin` content-type default. It writes the package to a temporary directory, reads it through `ctx.officeExchange.read`, edits one cell (`sheet-1`, row 1, column 1), and times only the `ctx.officeExchange.write` call. The 30,000 ms assertion is unchanged; the test carries a 120,000 ms timeout and removes its temporary directory in `finally`.

Host: Windows 11 (10.0.26100), Intel Core i7-14650HX, 32 GB RAM, Node v24.15.0, pnpm 9.15.9.

```sh
pnpm exec vitest run packages/document/office-exchange/tests/ooxml-preservation.spec.ts -t "large-workbook"
```

Outcome: 1 test passed. The measured save took **285 ms** for a **4,535,613-byte (4.33 MiB)** package holding **10,000 cells**; repeated runs on this host measured 285 ms, 345 ms, and 367 ms. The test asserts `elapsedMs < 30_000`.

## Gate results (2026-11-18)

Host: Windows 11 (10.0.26100), Node v24.15.0. Working tree is the feature tree, including uncommitted and untracked feature files. No stash, checkout, or tree mutation was used.

### 1. Focused per-file coverage — FAILED (exit 1)

```sh
pnpm exec vitest run packages/document/office-exchange/tests --coverage --coverage.include='packages/document/office-exchange/src/**/*.ts' --coverage.reportsDirectory=D:\kilo\temp\ooxml-final-coverage
```

- Test Files: 34 passed (34). Tests: 798 passed (798).
- Overall: statements 87.61% (7347/8386), branches 76.56% (4662/6089), functions 92.99% (1022/1099), lines 90% (6334/7037).
- Per-file 100% thresholds failed for seven files: `commands.ts`, `detect.ts`, `remote.ts`, `package/ooxml-package.ts`, `provider/index.ts`, `provider/docx.ts`, `provider/xlsx.ts`.
- Uncovered locations reported: 2543 (1039 statements, 1427 branch paths, 77 functions). Full `path:line:col` detail: `D:\kilo\temp\ooxml-coverage-final.txt`; HTML report: `D:\kilo\temp\ooxml-final-coverage`.

| File | % Stmts | % Branch | % Funcs | % Lines | Uncovered lines |
| --- | ---: | ---: | ---: | ---: | --- |
| src/commands.ts | 98.41 | 93.75 | 100 | 100 | 168 |
| src/detect.ts | 96.85 | 95.67 | 100 | 98.45 | 491-492, 579 |
| src/remote.ts | 98.02 | 97.18 | 100 | 97.89 | 201, 539-540, 886-889, 900, 1010 |
| src/package/ooxml-package.ts | 91.30 | 82.97 | 100 | 97.14 | 374, 383, 435-436, 501 |
| src/provider/docx.ts | 90.85 | 82.34 | 95.18 | 93.10 | 1003, 1063, 1096, 1251-1254, 1285, 1294, 1328, 1345, 1357, 1367, 1373, 1378, 1416, 1424, 1449, 1473, 1486, 1490, 1522, 1531, 1550, 1589, 1661, 1721, 1740, 1747, 1753, 1798, 1816, 1831, 1849-1850, 1892, 1921, 1940-1953, 1956-1974, 1993, 2056, 2094, 2125, 2135, 2139, 2146, 2176, 2192-2200, 2256-2259, 2262-2265, 2268-2271, 2274-2277, 2329-2349, 2411, 2421, 2446, 2477-2478, 2483-2499, 2538, 2544, 2555, 2572-2575, 2637, 2932, 3152, 3226-3231, 3279-3284, 3308, 3329, 3852, 3916, 4077-4078, 4107-4108, 4190, 4289, 4344-4357, 4376, 4379, 4386-4388, 4667, 5081, 5275-5276, 5689, 5697, 5844 |
| src/provider/index.ts | 93.45 | 88.09 | 100 | 93.81 | 214, 297, 310-313 |
| src/provider/xlsx.ts | 74.15 | 58.45 | 84.26 | 78.19 | 769-770, 782-783, 792-793, 862-863, 907, 920, 944, 973, 1297-1448, 1465-1526, 1620-1624, 1626-1628, 1752-1754, 1766-1768, 1794, 1917-1938, 1962-2067, 2080, 2090-2099, 2132, 2167-2179, 2193, 2213, 2237, 2259, 2335, 2702, 3014, 3047, 3055, 3085, 3125-3131, 3134, 3147, 3157, 3163, 3170, 3184, 3188, 3195, 3199, 3260, 3269-3272, 3280-3283, 3295, 3298, 3305, 3329, 3339, 3436, 3441, 3447, 3453, 3672, 3707, 3726, 3779, 3782, 3785, 3801-3806, 3823-3826, 3829-3832, 3841-3844, 3927-3930, 3933-3936, 3939-3942, 3960-3963, 3984, 4014-4015, 4069, 4086, 4101, 4137-4145, 4500, 4537, 4565-4567, 4613-4620, 4638-4655, 4666-4669, 4695-4696, 4713-4714, 4744-4803, 4809-4812, 4832, 4851, 4860, 4877, 4932-4939, 4944, 4953-4956, 4979-4982, 4991-4992, 5019-5033, 5053-5071, 5084-5086, 5107-5116, 5134-5136, 5141, 5163-5170, 5173, 5177-5180, 5184, 5217-5240, 5250-5251, 5255-5256, 5261-5268, 5309-5322, 5363, 5385, 5388, 5391, 5406-5410, 5422, 5425, 5428-5435, 5447, 5469-5478, 5544-5546, 5558-5560, 5615, 5625, 5630, 5637-5652, 5679-5690, 5693-5698, 5704-5707, 5724, 5760-5767, 5770-5778, 5802-5809, 5821, 5827, 5851, 5853, 5863, 5887, 5910, 5918, 5931, 5967, 5979, 5982, 5988, 6057, 6066-6070, 6083, 6153-6161, 6213, 6245, 6400, 6403, 6410, 6412-6413, 6435, 6442, 6472-6475, 6490-6491, 6496, 6506-6507, 6516 |

Classification: (a) primarily caused by this feature's files. Six of the seven failing files are modified or untracked feature files (`detect.ts`, `remote.ts`, `provider/index.ts`, `docx.ts`, `xlsx.ts`, and new `package/ooxml-package.ts`), and most uncovered locations fall inside the feature's added hunks (`docx.ts` added block 922-2448; `xlsx.ts` added block 2715-6596). One sub-gap is (b): `commands.ts` is unmodified at HEAD and its uncovered line 168 pre-exists this feature. `xlsx.ts` also reports gaps below 2715 on unchanged lines; no HEAD baseline coverage run was performed, so those older-region gaps are not separately attributed.

### 2. `pnpm run typecheck` — FAILED (exit 1)

- `run-gates`: 3 passed, 2 failed. Failed gates: host aggregate check (`tsc -b tsconfig.host.json`), client aggregate check (`tsc -b tsconfig.client.json`).
- 18 TS errors in six feature test files: `packages/client/ui-office-editor/tests/office-roundtrip.spec.tsx` (1), `packages/document/office-exchange/tests/apply-splices.spec.ts` (6), `ooxml-preservation.spec.ts` (1), `provider-docx.spec.ts` (3), `provider-xlsx.spec.ts` (5), `xlsx-splice.spec.ts` (2). First errors:

```text
office-roundtrip.spec.tsx(80,5): error TS2322: Type '{ ... }' is not assignable to type 'JsonValue'.
apply-splices.spec.ts(79,54): error TS2345: Argument of type 'Uint8Array<ArrayBuffer>' is not assignable to parameter of type 'string'.
provider-docx.spec.ts(584,48): error TS2345: Type '{ maxFileBytes: number; maxCells: number; maxConversionMs: number; commandTimeoutMs: number; }' is not assignable to parameter of type 'OfficeExchangeConfig'.
  Type ... is missing the following properties from type 'OfficeExchangeConfig': maxSourceAnchors, maxSourceAnchorBytes
xlsx-splice.spec.ts(307,54): error TS2532: Object is possibly 'undefined'.
```

Classification: (a). `provider-docx.spec.ts` and `provider-xlsx.spec.ts` are unmodified at HEAD, but the feature adds required `maxSourceAnchors` and `maxSourceAnchorBytes` fields in `src/config.ts`, so their errors are feature-caused. The other four files are modified or untracked feature tests.

### 3. `pnpm run lint` — FAILED (exit 1)

- 30 errors and 1 warning. Grouped by file: `src/provider/xlsx.ts` 15; `credentials-local/tests/drain.spec.ts` 6; `tests/write-postcondition.spec.ts` 5; `tests/document-tools.spec.ts` 2; `tests/remote-edges.spec.ts` 1; `scripts/verify-forwork-identity.spec.ts` 1.
- Representative first diagnostics:

```text
packages/document/office-exchange/src/provider/xlsx.ts:6096:36: error @stylistic(arrow-parens)
packages/credentials/credentials-local/tests/drain.spec.ts:44:9: error typescript(await-thenable)
packages/credentials/credentials-local/tests/drain.spec.ts:44:9: error typescript(no-confusing-void-expression)
packages/document/office-exchange/tests/remote-edges.spec.ts:450:33: error typescript(no-unsafe-assignment)
packages/document/office-exchange/tests/write-postcondition.spec.ts:161:32: error typescript(no-unnecessary-type-assertion)
packages/document/tool-office/tests/document-tools.spec.ts:792:33: error typescript(no-base-to-string)
scripts/verify-forwork-identity.spec.ts:36:24: error typescript(no-confusing-void-expression)
packages/document/office-exchange/src/provider/xlsx.ts:1212:29: error typescript(no-unnecessary-type-conversion)
packages/document/office-exchange/src/provider/xlsx.ts:1519:29: error typescript(no-unsafe-return)
packages/document/office-exchange/src/provider/xlsx.ts:2341:28: error typescript(no-unnecessary-condition)
```

Classification: mixed. (a) 23 errors in feature files: `xlsx.ts` 15, `write-postcondition.spec.ts` 5, `document-tools.spec.ts` 2, `remote-edges.spec.ts` 1. (b) 7 pre-existing errors in files unmodified at HEAD: `credentials-local/tests/drain.spec.ts` 6 and `scripts/verify-forwork-identity.spec.ts` 1. The single warning is a pre-existing unused oxlint-disable directive in `packages/experimental/client-ui-cot-translation/src/index.ts` (unmodified).

### 4. `pnpm run build:artifacts` — PASSED (exit 0)

`run-gates`: 8 passed, 0 failed, 0 skipped in 117.55s. `build: recorded 524 client artifact(s) with 3 public value(s)`.

### 5. `pnpm run build:check` — FAILED (exit 1)

- `run-gates`: 0 passed, 2 failed in 56.64s: host aggregate check (`tsc -p tsconfig.host.json`) and client aggregate check (`tsc -p tsconfig.client.json`).
- Same 18 TS errors as `pnpm run typecheck` (first: `office-roundtrip.spec.tsx(80,5) TS2322`; `apply-splices.spec.ts(79,54) TS2345`; `provider-docx.spec.ts(584,48) TS2345` missing `maxSourceAnchors`/`maxSourceAnchorBytes`).

Classification: (a), the same feature test and config-contract errors.

### 6. `pnpm run hygiene` — PASSED (exit 0)

`run-gates`: 26 passed, 0 failed, 0 skipped in 304.00s.

### 7. `pnpm run doc-sync` — FAILED (exit 1)

`run-gates`: 40 passed, 3 failed, 0 skipped in 174.52s.

- **CLI help reference** — (c) environmental. The spawned `node --import tsx/esm apps/cli/src/bin.ts acp --help` was killed (`code: null`, `signal: 'SIGKILL'`, empty stdout/stderr). Standalone rerun `pnpm run verify-cli-help` exited 0: `gen-cli-help: 7 help pages; all outputs are current.` No new documentation failure.
- **Config catalog** — (b) pre-existing at HEAD, matching the known `tool-office` inject failure. `scripts/gen-config-catalog.ts --check` reports `docs/config-catalog.md` and `.zh.md` stale. Evidence: both files are unmodified in the working tree, and HEAD source `packages/document/tool-office/src/index.ts` declares `inject = ['agents', 'tools', 'officeDocument']` while committed `docs/config-catalog.md:4245` says `tools · officeDocument` (missing `agents`).
- **Skill plugin package list** — (b) pre-existing at HEAD. `scripts/gen-plugin-packages.ts --check` reports `packages/preset/agent-preset/skills/cordis-composition-reference/references/packages.md` stale; that generated file is unmodified in the working tree.

No new doc-sync failure was attributed to the feature.

### 8. `pnpm run duplication` — FAILED (exit 1)

- 27 TypeScript clones: 296 duplicated lines (0.05%) and 2096 duplicated tokens (0.07%) across 2616 files / 551208 lines. `tsx`: 0 clones.
- Clone locations: `ui-office-editor` `editor/docs.ts` (2 pairs), `editor/sheets.ts`, `useOfficeDocument.ts` (1 pair); `office-exchange` `package/zip-format.ts` (1 pair), `provider/docx.ts` (6 pairs), `provider/xlsx.ts` (3 pairs), `remote.ts` (2 pairs), `source/splice.ts` (10 pairs), `source/xml-spans.ts` (1 pair); `tool-office/src/edit-cells.ts` (1 pair).
- Classification: (a) for 26 clones, all involving modified or untracked feature files (`zip-format.ts`, `splice.ts`, and `xml-spans.ts` are untracked additions). (b) for the single `tool-office/src/edit-cells.ts` pair at 176-184 / 207-215: that file is unmodified at HEAD and already contains both duplicated `inverseRow`/`inverseColumn` bodies.

### Summary

| Gate | Exit | Result |
| --- | ---: | --- |
| Focused office-exchange coverage | 1 | FAILED — 798 tests pass; per-file 100% thresholds fail |
| `pnpm run typecheck` | 1 | FAILED — 18 TS errors |
| `pnpm run lint` | 1 | FAILED — 30 errors (23 feature, 7 pre-existing) |
| `pnpm run build:artifacts` | 0 | PASSED |
| `pnpm run build:check` | 1 | FAILED — same 18 TS errors |
| `pnpm run hygiene` | 0 | PASSED |
| `pnpm run doc-sync` | 1 | FAILED — 3 gates: 1 environmental, 2 known pre-existing |
| `pnpm run duplication` | 1 | FAILED — 27 clones (26 feature, 1 pre-existing) |

## Final gate results (2026-11-18)

Re-run from the repository root on Windows 11 (Node v24.15.0) after the latest working-tree changes.

### 1. Focused office-exchange coverage — FAILED (exit 1)

```sh
pnpm exec vitest run packages/document/office-exchange/tests --coverage --coverage.include='packages/document/office-exchange/src/**/*.ts' --coverage.reportsDirectory=D:\kilo\temp\ooxml-cov-final2 --testTimeout=20000
```

- Test files 47 passed (47), tests 1108 passed (1108), duration 45.39s.
- Aggregate: statements 99.98% (7597/7598), branches 99.98% (5206/5207), functions 100% (1061/1061), lines 100% (6480/6480).
- Files below 100%: only `packages/document/office-exchange/src/commands.ts` — statements 98.41%, branches 93.75%, functions 100%, lines 100%. Uncovered: `commands.ts:168:5` branch (`if`, path 1/2) and `commands.ts:168:38` statement — the `if (registrations === undefined) return` early return in `detachTarget`.
- Classification: (b) pre-existing at HEAD. `git status --porcelain packages/document/office-exchange/src/commands.ts` is empty (unmodified). HEAD's tracked `tests/commands.spec.ts` also never detaches an unregistered document, so the line was already uncovered; the expanded suite only surfaces it.

### 2. `pnpm run typecheck` — FAILED (exit 1)

`run-gates`: 4 passed, 1 failed; host aggregate check (`tsc -b tsconfig.host.json`) exit 1 with 66 errors, all in untracked office-exchange test files:

| File | Errors |
| --- | ---: |
| `tests/docx-write-coverage.spec.ts` | 16 |
| `tests/docx-splice-coverage.spec.ts` | 14 |
| `tests/xlsx-splice.spec.ts` | 10 |
| `tests/remote-coverage.spec.ts` | 8 |
| `tests/save-disposition.spec.ts` | 6 |
| `tests/xlsx-resources.spec.ts` | 4 |
| `tests/provider-index-coverage.spec.ts` | 2 |
| `tests/xlsx-cell-mapping.spec.ts` | 2 |
| `tests/xlsx-source-state.spec.ts` | 2 |
| `tests/detect-coverage.spec.ts` | 1 |
| `tests/docx-read-coverage.spec.ts` | 1 |

Every implicated file is untracked (`git status --porcelain <path>` returns `?? <path>`), so none exists at HEAD. Classification: (a), type mismatches in this feature's new test files (`SpliceOperation` vs `ByteSpliceOperation`, `CellHyperlinkValue.text`, Univer `Anchor`/`IBorderStyleData.cl`/`ITextRotation.a`, missing `OfficeExchangeConfig` fields, `DateSystem`).

### 3. `pnpm run lint` — FAILED (exit 1)

7 errors and 1 warning; no feature file is reported:

- `packages/credentials/credentials-local/tests/drain.spec.ts`: 6 errors (await-thenable + no-confusing-void-expression at 44, 58, 89).
- `scripts/verify-forwork-identity.spec.ts:36:24`: 1 error (no-confusing-void-expression).
- `packages/experimental/client-ui-cot-translation/src/index.ts:82:9`: warning, unused oxlint-disable directive.

`git status --porcelain` is empty for all three paths. Classification: (b) pre-existing at HEAD (the earlier 23 feature lint errors are gone).

### 4. `pnpm run build:check` — FAILED (exit 1)

`run-gates`: 1 passed, 1 failed; host aggregate check (`tsc -p tsconfig.host.json`) exit 1 with the same 66 errors in the same 11 untracked test files, no additional errors. Classification: (a), same feature test files as gate 2.

### 5. `pnpm run duplication` — FAILED (exit 1)

One clone: `packages/document/tool-office/src/edit-cells.ts [176:23-184:35]` vs `[207:29-215:35]` (9 lines, 74 tokens); totals 1 clone, 8 duplicated lines (0.00%), 74 duplicated tokens (0.00%) across 3050 files. `git status --porcelain packages/document/tool-office/src/edit-cells.ts` is empty. Classification: (b) pre-existing at HEAD (the file was last changed by the committed T125 work); the earlier 26 feature clones are resolved.

### 6. `pnpm run doc-sync` — FAILED (exit 1)

`run-gates`: 40 passed, 3 failed, 0 skipped in 179.58s.

- **CLI help reference** — (c) environmental. In-gate child `node --import tsx/esm apps/cli/src/bin.ts acp --help` was killed (`code: null`, `killed: true`). Standalone rerun `pnpm run verify-cli-help` exited 0: `gen-cli-help: 7 help pages; all outputs are current.`
- **Config catalog** — (b) pre-existing at HEAD. `docs/config-catalog.md` and `.zh.md` are stale; computed diff is only the `@forwork/tool-office` row (generated `inject: agents · tools · officeDocument`, pointer `index.ts:41`; committed `tools · officeDocument`, pointer `index.ts:40`). Both catalog files are unmodified and `packages/document/tool-office/src/index.ts` already declares `agents` at HEAD; the catalog was last regenerated on 2026-11-17, before the later T125 change on 2026-11-18.
- **Skill plugin package list** — (b) pre-existing at HEAD. `packages/preset/agent-preset/skills/cordis-composition-reference/references/packages.md` line 74 is stale for `@forwork/client-ui-brand-official` (generated `Official ForWork sidebar brand artwork`, committed `Official ForWork brand artwork and click-triggered home background`). The generated file and `packages/client/ui-brand-official/package.json` are unmodified; the description changed in T120 on 2026-11-17 after the last regeneration on 2026-11-16.

No doc-sync failure was attributed to the feature.

### 7. Combined four-package suite — PASSED (exit 0)

```sh
pnpm exec vitest run packages/document/office-exchange/tests packages/client/ui-office-editor/tests packages/document/tool-office/tests packages/session/session-document/tests
```

Test files 82 passed (82), tests 1547 passed (1547), duration 36.65s.

### Final summary

| # | Command | Exit | Result | Classification |
| ---: | --- | ---: | --- | --- |
| 1 | Focused office-exchange coverage | 1 | 1108 tests pass; `src/commands.ts` line 168 at 98.41% stmts / 93.75% branches; aggregate 99.98% stmts, 99.98% branches, 100% functions, 100% lines | (b) pre-existing at HEAD — `git status --porcelain packages/document/office-exchange/src/commands.ts` empty |
| 2 | `pnpm run typecheck` | 1 | 66 TS errors in 11 untracked office-exchange test spec files | (a) feature — all implicated paths are untracked (`??`) |
| 3 | `pnpm run lint` | 1 | 7 errors + 1 warning in `credentials-local/tests/drain.spec.ts`, `scripts/verify-forwork-identity.spec.ts`, `client-ui-cot-translation/src/index.ts` | (b) pre-existing at HEAD — `git status --porcelain` empty for all three |
| 4 | `pnpm run build:check` | 1 | Same 66 TS errors; host aggregate check fails | (a) feature — same untracked spec files as gate 2 |
| 5 | `pnpm run duplication` | 1 | 1 clone inside `tool-office/src/edit-cells.ts` (176-184 / 207-215); 8 lines, 74 tokens | (b) pre-existing at HEAD — `git status --porcelain packages/document/tool-office/src/edit-cells.ts` empty |
| 6 | `pnpm run doc-sync` | 1 | 40 passed, 3 failed: CLI help reference (killed child, passes standalone), config catalog stale (`tool-office` inject), plugin package list stale (brand description) | (c) environmental for CLI help; (b) pre-existing at HEAD for both stale catalogs — generated and input files unmodified |
| 7 | Combined four-package vitest run | 0 | 82 files, 1547 tests passed | PASSED |

No grade is claimed as passing where the command exited non-zero.

## Post-fix status (2026-11-18, final)

Appended after the earlier sections without editing them. Host: Windows 11 (10.0.26100), Node v24.15.0, pnpm 9.15.9. Working tree includes uncommitted and untracked feature files.

| # | Command | Exit | Summary |
| ---: | --- | ---: | --- |
| 1 | `pnpm run typecheck` | 0 | 5 passed, 0 failed (36.16s) |
| 2 | `pnpm run build:check` | 0 | 2 passed, 0 failed (23.32s) |
| 3 | `pnpm run lint` | 1 | 7 errors, 1 warning; all in unmodified files |
| 4 | `pnpm run duplication` | 0 | 0 clones across 3050 files / 622789 lines |
| 5 | `pnpm run doc-sync` | 1 | 39 passed, 4 failed (177.50s) |
| 6 | Combined four-package vitest run | 0 | 82 files, 1547 tests passed (38.45s) |
| 7 | Focused office-exchange coverage | 1 | 1108 tests passed; 2 files below 100% |

Command 6: `pnpm exec vitest run packages/document/office-exchange/tests packages/client/ui-office-editor/tests packages/document/tool-office/tests packages/session/session-document/tests`.

Command 7: `pnpm exec vitest run packages/document/office-exchange/tests --coverage --coverage.include='packages/document/office-exchange/src/**/*.ts' --coverage.reportsDirectory=D:\kilo\temp\ooxml-cov-final3 --testTimeout=20000`. Aggregate: statements 99.98% (7597/7598), branches 99.96% (5205/5207), functions 100% (1061/1061), lines 100% (6480/6480). Files below 100%:

| File | % Stmts | % Branch | % Funcs | % Lines | Uncovered |
| --- | ---: | ---: | ---: | ---: | --- |
| `src/commands.ts` | 98.41 | 93.75 | 100 | 100 | 168 |
| `src/provider/xlsx.ts` | 100 | 99.94 | 100 | 100 | 5316 |

### Non-zero failures

- **`pnpm run lint` — (b) pre-existing at HEAD.** 7 errors and 1 warning: `packages/credentials/credentials-local/tests/drain.spec.ts` 6, `scripts/verify-forwork-identity.spec.ts` 1, unused-directive warning in `packages/experimental/client-ui-cot-translation/src/index.ts`. `git status --porcelain` and `git diff --stat` are empty for all three paths; no feature file is reported.
- **doc-sync CLI help — (c) environmental.** In-gate child `node --import tsx/esm apps/cli/src/bin.ts acp --help` was killed (`code: null`, `killed: true`, `signal: 'SIGKILL'`, empty stdout/stderr). Standalone `pnpm run verify-cli-help` exited 0: `gen-cli-help: 7 help pages; all outputs are current.`
- **doc-sync config catalog — (b) pre-existing at HEAD.** Standalone `pnpm run verify-config-catalog` exits 1: `docs/config-catalog.md` and `.zh.md` stale. Both catalog files and `packages/document/tool-office/src/index.ts` have empty `git status --porcelain`; no feature file is involved.
- **doc-sync skill plugin package list — (b) pre-existing at HEAD.** Standalone `pnpm run verify-plugin-packages` exits 1 for `packages/preset/agent-preset/skills/cordis-composition-reference/references/packages.md`. That file and `packages/client/ui-brand-official/package.json` have empty `git status --porcelain`.
- **doc-sync repository references — pre-existing content in this file's earlier section.** Standalone `pnpm run verify-repository-references` exits 1 on `specs/002-in-place-ooxml-save/quickstart-results.md:274`, `:281`, `:282` (`commit-hash`). Those lines are in the earlier "Final gate results" section; no line in this appended section is reported. Earlier sections were out of scope.
- **Coverage `src/commands.ts` — (b) pre-existing at HEAD.** The file is unmodified (`git status --porcelain` empty); the same line 168 early-return gap was recorded in the earlier section.
- **Coverage `src/provider/xlsx.ts` — (a) feature.** The file is modified (+4460/-70); the line 5316 `color ?? '#000000'` branch does not exist at HEAD and was covered in the earlier run. Not pre-existing and not environmental.

### Earlier environment evidence

- LibreOffice load smoke: the earlier T033 report ("Add LibreOffice load smoke") ran `pnpm exec vitest run packages/document/office-exchange/tests/ooxml-engine.spec.ts` on this host and recorded 2 passed (1 file, 38.41s) with the engine available, so the render path executed rather than skipping.
- Web document-preview save: the earlier T034 report and the "Fix editor removeChild on save" report ran `DSH_SNAPSHOT=refresh pnpm exec vitest run --config vitest.web.config.ts apps/web/tests/document-preview.e2e.ts -t "opens text, isolated HTML"` and the same command with `DSH_SNAPSHOT=replay`; each recorded 1 passed, 6 skipped, with the save assertions passing and the snapshot line `- DOCX save: completed without error; 1 preserved media part unchanged`.

No non-zero command above is reported as passing.

## Coverage closure and final corrections (2026-11-18)

After the post-fix run:

- **Coverage `src/provider/xlsx.ts` — now 100%.** A border-edge test (thin top edge with no color, hair bottom edge with `#112233`) exercises both arms of `color ?? '#000000'`; focused coverage reports statements 100% (2341/2341), branches 100% (1871/1871), functions 100% (347/347), lines 100% (2034/2034).
- **Full office-exchange suite**: 47 files, 1109 tests passed; `pnpm exec tsc -b packages/document/office-exchange` clean.
- **Repository references gate — now passes.** The three bare commit hashes in the earlier "Final gate results" section (`quickstart-results.md:274`, `:281`, `:282`) were replaced with task/date references; `pnpm run doc-sync` reports `PASS repository references (91.05s)`.
- **`pnpm run doc-sync`**: 40 passed, 3 failed, 0 skipped (183.90s). Remaining failures are the two pre-existing stale generated catalogs (config catalog for `@forwork/tool-office` inject; skill plugin package list for `@forwork/client-ui-brand-official`) and the environmental CLI-help SIGKILL that passes standalone.

Final gate status for the feature:

| Gate | Exit | Status |
| --- | ---: | --- |
| `pnpm run typecheck` | 0 | PASS |
| `pnpm run build:check` | 0 | PASS |
| `pnpm run duplication` | 0 | PASS (0 clones) |
| Feature per-file coverage | — | PASS — every modified/new `src` file at 100%; only unmodified `src/commands.ts:168` remains below 100% (pre-existing) |
| `pnpm run lint` | 1 | 7 errors + 1 warning, all in unmodified files (pre-existing) |
| `pnpm run doc-sync` | 1 | 2 pre-existing stale catalogs + 1 environmental CLI-help SIGKILL; no feature-caused failure |
| LibreOffice load smoke | 0 | 2 passed |
| Web document-preview save | 0 | refresh and replay passed |
| Manual Word/Excel matrix | — | NOT RUN — requires a human with desktop Word and Excel (T061) |

## Review remediation (2026-11-18)

Appended after the earlier sections without editing them. A branch review produced 20 findings against the in-place write path, its client save surface, and the documentation. The fixes below, including the three formerly residual writer items, are in the working tree and verified by the final run in this section.

Evidence reported by the review tracks: `packages/document/office-exchange/tests` passed 1,147 tests; the combined four-package suite (`office-exchange`, `ui-office-editor`, `tool-office`, `session-document`) passed 1,428 tests before the residual fixes; focused coverage reported 100% statements, branches, functions, and lines for the touched source files.

### Findings and fixes

| # | Finding | Fix |
| ---: | --- | --- |
| 1 | Word root metadata (`w:rsid*`, `w14:paraId`, and `w14:textId`) was classified as an unreproduced coupled property, so an ordinary paragraph or run edit could refuse on metadata the editor never models. | The independent-property classifier returns Word root metadata separately from coupled candidates and re-emits it on the rendered root tag. |
| 2 | A stale sidecar whose digest or length differed from the supplied source was validated against inferred part lengths before re-anchoring, so ordering could reject a legitimate second save. | A mismatched sidecar is validated structurally first; the writer then re-anchors from the supplied source before planning. |
| 3 | Re-anchored spans were not aligned with the current model units, so a stale sidecar could pair a rebuilt anchor with the wrong unit. | Re-anchor plans align rebuilt anchors to the current model units in source order before the splice plan consumes them. |
| 4 | The span scanners mixed UTF-16 string indices with UTF-8 byte offsets, so multibyte text, a BOM, or a CRLF pair could shift a splice. | The scanners report source byte offsets only; string positions are used only where the XML parser reports character positions. |
| 5 | A unit span's string position and its byte position could disagree after decoding, so a splice could start inside a multibyte character. | Span starts and ends come from the byte scanner; decoded strings are sliced only for content comparisons. |
| 6 | A newly appended `w:abstractNum` could be placed after the first existing `w:num`, violating numbering schema order. | New abstracts insert before the first `w:num`, and new `w:num` instances insert before `</w:numbering>`. |
| 7 | Preserved construct tags were listed in more than one detection path and could drift apart. | `detect.ts` owns the tag sets, and every detection caller classifies through that single source. |
| 8 | A style-only edit on a cell backed by a rich-text shared string was flattened instead of refused, dropping its runs. | The rich-text shared-string guard applies to every edit of that cell, including a style-only change. |
| 9 | A worksheet row or column shift could leave merged, conditional-formatting, or data-validation ranges at stale coordinates. | A structural shift on a sheet carrying one of those address ranges refuses `unsupported-feature` and names the construct. |
| 10 | A string value whose model claimed a numeric cell type was written without XML escaping. | A string value is escaped even when the model type says numeric. |
| 11 | ZIP entry inflation trusted the declared uncompressed size. | The declared size is bounded before inflation, `maxOutputLength` caps the stream, and the inflated length and CRC are checked against the header. |
| 12 | Shared-string classification restarted its scan for each entry. | One forward cursor walks `xl/sharedStrings.xml` linearly and records each entry's rich-text flag. |
| 13 | Row alignment rescanned candidate model rows for every source row. | Model rows are indexed by row number in one map, and the alignment pass uses map lookups. |
| 14 | An open inflated and scanned the package once for preserved constructs and again for unsupported features. | `read` classifies both reports from one `detectOfficePackage` scan. |
| 15 | Many inserted rows or cells caused repeated ordering or offset work. | Worksheet operations are sorted once and applied in one forward pass, with later insertions offset from the preceding result. |
| 16 | The span walker and the structure walker duplicated markup-declaration handling. | Both walkers share the `declarationEnd` helper and the same markup-declaration rules. |
| 17 | The Client kept a dead `unsupported` list and a read-side save gate after preserved constructs stopped blocking saves. | The dead list and gate are removed; the panel lists preserved-not-editable classes and keeps Save enabled, and a refusal remains edit-scoped. |
| 18 | Disposition `names` groups were not capped consistently. | Every group is sorted ASCII and truncated to `SAVE_DISPOSITION_NAME_LIMIT` (256), with `nameLimit` reported on the disposition. |
| 19 | Dead span bookkeeping remained in the splice path after the row-level and cell-level plans were unified. | The unused span bookkeeping is removed; row-level parts use row spans and cell-level parts use cell spans only. |
| 20 | The package API inflated a source part only to discover that the caller's replacement bytes differ. | `replacePartKnownDifferent` records a replacement without inflating the source part, and a regression test proves an unreadable source part still takes that path. |

### Residual fixes closed

The three residual writer items are resolved in the working tree.

- **Known-different call sites:** `src/provider/xlsx.ts:6813` calls `replacePartKnownDifferent` after `applyWorksheetOperations` has compared the spliced bytes with the source and skipped an equal result. `src/provider/docx.ts:2832` keeps the equality-checking `replacePart` because that writer has no prior byte comparison, so a no-op DOCX save still returns the source package unchanged. `xlsx-splice.spec.ts:523` spies on both package methods and asserts a no-op save calls neither while a changed worksheet calls only the known-different path.
- **Row-only cell `r`:** `src/provider/xlsx.ts:6492` rewrites a paired cell's `r` when either its row or its column differs from the source address. `xlsx-splice.spec.ts:798` asserts every output `<row>`/`<c r>` address and reopened value after a row insert on a plain sheet; `:2540` covers a top-row insert on a non-refused vertical-merge sheet (`A1` becomes `A2`/`A3`).
- **autoFilter and Excel-table structural guard:** `autoFilter` is in `WORKSHEET_RANGE_KINDS` (`src/provider/xlsx.ts:174`), and the worksheet scan collects `tablePart` relationship ids (`:3978`) that `worksheetTableRanges` (`:4068`) resolves from `xl/worksheets/_rels/sheetN.xml.rels` and the table part's `ref`. Both feed the structural refusal at `:3324`, which names the construct and range (for example `auto filter range A1:B2` or `table range A1:B3`). `xlsx-splice.spec.ts:2452` covers autoFilter and `:2472` covers an Excel table; plain-sheet inserts and deletes still save.

A malformed worksheet whose `tablePart` relationship cannot be resolved, or whose table part carries no `ref`, contributes no range and therefore does not trigger the refusal. Valid OOXML tables always resolve.

### Final verification (2026-11-19)

Run on Windows 11 with Node v24.15.0 and the uncommitted working tree. The clean main worktree at `H:/ForWork` at this worktree's HEAD commit is the baseline for pre-existing classifications; its `git status --porcelain` was empty.

| # | Command | Exit | Summary | Classification |
| ---: | --- | ---: | --- | --- |
| 1 | `pnpm run gen-cordis-catalog` | 0 | 131 artifacts computed, 1 written: `packages/extensions/tool-cordis/src/api-catalog.ts`; `pnpm run verify-cordis-catalog` then exited 0 with all 131 artifacts current | Feature regeneration; now current |
| 2 | `pnpm run doc-sync` | 1 | Final rerun after this section's edit: 41 passed, 2 failed, 0 skipped (132.08s), failing config catalog and skill plugin package list. The earlier run was 40 passed, 3 failed (220.21s) with the CLI help child SIGKILLed. | All non-feature; details below |
| 3 | `pnpm exec vitest run packages/document/office-exchange/tests packages/client/ui-office-editor/tests packages/document/tool-office/tests packages/session/session-document/tests` | 1 | 83 files: 82 passed, 1 failed; 1,585 tests passed, 5 failed, plus 1 unhandled rejection | Feature-caused `tool-office` expectation mismatch |
| 4 | `pnpm run typecheck` | 0 | 5 passed, 0 failed (97.46s) | PASS |
| 5 | `pnpm run build:check` | 0 | 2 passed, 0 failed (12.92s) | PASS |
| 6 | `pnpm run lint` | 1 | 9 errors, 1 warning: 7 errors and the warning in unmodified files; 2 feature errors in `office-exchange/src/provider/index.ts` | Mixed: pre-existing + feature |
| 7 | `pnpm run duplication` | 1 | 1 clone in `office-exchange/src/package/ooxml-package.ts` (13 lines, 93 tokens); 0 clones elsewhere across 3,050 files | Feature-caused |
| 8 | `pnpm exec vitest run packages/document/office-exchange/tests --coverage --coverage.include='packages/document/office-exchange/src/**/*.ts' --coverage.reportsDirectory=D:\kilo\temp\ooxml-review-final --testTimeout=20000` | 1 | 48 files, 1,152 tests passed; aggregate statements 99.98% (7,850/7,851), branches 99.98% (5,310/5,311), functions 100% (1,095/1,095), lines 100% (6,698/6,698) | One pre-existing file below 100% |
| 9 | `pnpm exec vitest run packages/document/office-exchange/tests/ooxml-engine.spec.ts` | 0 | 1 file, 2 tests passed (6.11s); LibreOffice engine available | PASS |
| 10 | `DSH_SNAPSHOT=replay pnpm exec vitest run --config vitest.web.config.ts apps/web/tests/document-preview.e2e.ts -t "opens text, isolated HTML"` | 0 | 1 passed, 6 skipped (61.80s); replay matched the committed snapshot, so no refresh was needed | PASS |

Command 1 net diff in `api-catalog.ts`: the `read`/`save`/`detectUnsupported` Remote entries and the `OfficeUnitProvider.write` entry now carry the source-package, disposition, and coupled-refusal contracts; `ExchangeReadResult`, `OfficeDocumentReadResult`, `OfficeDocumentSaveResult`, `OfficeExtension`, and `OfficeFormat` gain the macro formats, `preserved`, and `disposition`; and `OfficeWriteResult`, `SaveDisposition`, `SaveDispositionGroup`, `SourcePackage`, `PreservedConstructKind`, `ZipArchive`, `ZipEntryRecord`, and `ZipMethod` are added. `docs/subsystems/office-exchange.md` and `.zh.md` were already current and the generator did not rewrite them this run; their generated regions and the pairing manifest are covered by the 131-artifact check.

Command 3 failure: `packages/document/tool-office/tests/document-tools.spec.ts` has 5 failed tests in the combined run. The deterministic failures are `deletes worksheet rows and restores their cells in the inverse`, `inserts worksheet columns and records a structural inverse`, `deletes worksheet columns and restores their cells in the inverse`, and `applies a mixed batch through one command`; each asserts `status: ok` while the save now refuses with `unsupported-feature` because the `roundtrip.xlsx` fixture carries `mergeCell ref="A7:C7"` and the new guard refuses any row or column shift on a sheet carrying an address range (`merged range A7:C7 ... cannot follow an insert, delete, or move of rows or cells`). The failing spec alone reports 4 failed, 45 passed. The fifth combined-run failure, `applies an edit through a registered target and saves it once`, passed when the spec ran alone; the combined run also recorded an unhandled `RemoteError: "roundtrip.xlsx" changed while it was being read.` These tests were not updated for the new refusal, so the gate is reported as failed, not passed.

Command 6 details: feature errors are `packages/document/office-exchange/src/provider/index.ts:165:38` and `:179:38`, both `typescript(only-throw-error)` on `throw macroRegenerationRefusal(format)`; that helper is `never` but oxlint does not narrow a call expression. Pre-existing are six `await-thenable`/`no-confusing-void-expression` errors in `packages/credentials/credentials-local/tests/drain.spec.ts`, one `no-confusing-void-expression` error in `scripts/verify-forwork-identity.spec.ts`, and an unused-disable warning in `packages/experimental/client-ui-cot-translation/src/index.ts`; `git status --porcelain` is empty for all three paths.

Command 7 clone: `packages/document/office-exchange/src/package/ooxml-package.ts` [129:14–142:6] duplicates [161:28–174:6] across `replacePart` and `replacePartKnownDifferent`. The file is new in this feature.

Command 8 file below 100%: `packages/document/office-exchange/src/commands.ts` at statements 98.41% (62/63), branches 93.75% (15/16), functions 100%, lines 100%; the gap is line 168 `if (registrations === undefined) return` in `detachTarget`. The file is unmodified by this feature (`git status --porcelain` empty), and the clean main baseline reports the same 98.41%/93.75% for the same line. Every other file under `src` is at 100%.

doc-sync classifications:

- **config catalog — pre-existing (b).** `verify-config-catalog` reports `docs/config-catalog.md` and `.zh.md` stale. The generated difference is `@forwork/tool-office`: `inject` gains `agents` and the source pointer moves from `packages/document/tool-office/src/index.ts:40` to `:41`. `git status --porcelain` is empty for both catalogs and for `packages/document/tool-office/src/index.ts`; the same two files are stale on clean main.
- **skill plugin package list — pre-existing (b).** `verify-plugin-packages` reports `packages/preset/agent-preset/skills/cordis-composition-reference/references/packages.md` stale at the `@forwork/client-ui-brand-official` row: committed description `Official ForWork brand artwork and click-triggered home background` versus generated `Official ForWork sidebar brand artwork`. The list file and `packages/client/ui-brand-official/package.json` are unmodified, and clean main produces the same row.
- **CLI help reference — environmental (c), intermittent.** The first gate run's child `node --import tsx/esm apps/cli/src/bin.ts web --help` (and `acp --help`) printed the correct usage to stdout and was then killed with `code: null`, `killed: true`, `signal: 'SIGKILL'` after 110.64s; standalone `pnpm run verify-cli-help` here and on clean main also exited 1 with the same SIGKILL. The final doc-sync rerun passed `CLI help reference` (93.43s), so the kill depends on host child-process timing under load. `apps/cli` is unmodified and no feature source is involved.

Remaining non-feature doc-sync failures in the final run are exactly the config catalog and the skill plugin package list. The CLI help reference is intermittently SIGKILLed but is not failing in the final run. No doc-sync failure is feature-caused. Feature-caused failures elsewhere are the 5 `tool-office` test failures, the 2 `only-throw-error` lint errors, the 1 duplication clone, and none in coverage, LibreOffice, Web replay, typecheck, or build:check.

## Post-review closure (2026-11-19)

This section supersedes rows 3, 6, and 7 of the prior final table. After the tool-test, lint, and duplication fixes:

| # | Command | Exit | Summary | Classification |
| ---: | --- | ---: | --- | --- |
| 1 | `pnpm exec vitest run packages/document/office-exchange/tests packages/client/ui-office-editor/tests packages/document/tool-office/tests packages/session/session-document/tests` | 0 | 83 files, **1,591 tests passed** | PASS |
| 2 | `pnpm run typecheck` | 0 | 5/5 gates passed | PASS |
| 3 | `pnpm run build:check` | 0 | 2/2 passed | PASS |
| 4 | `pnpm run lint` | 1 | Feature files clean; 7 errors + 1 warning remain only in unmodified files | Pre-existing only |
| 5 | `pnpm run duplication` | 0 | **0 clones** across 3,050 files | PASS |
| 6 | office-exchange suite | 0 | 48 files, **1,152 tests passed** | PASS |
| 7 | Focused coverage (`src/**/*.ts`) | 1 | Every modified/new file at 100%; only unmodified `src/commands.ts:168` at 98.41%/93.75% | Pre-existing only |
| 8 | `ooxml-engine.spec.ts` | 0 | 2 passed (LibreOffice available) | PASS |
| 9 | Web document-preview replay | 0 | 1 passed, 6 skipped; snapshot matched | PASS |
| 10 | `pnpm run verify-cordis-catalog` | 0 | 131 artifacts current after regeneration | PASS |
| 11 | `pnpm run doc-sync` | 1 | 41 passed / 2 failed: config catalog and skill plugin package list | Pre-existing only |

### Final review-fix notes

- The four `tool-office` structural tests now run on a merge-free copy and still prove row/column insert/delete; a new case asserts that the untouched merged fixture refuses with `merged range A7:C7`, writes nothing, and records no fact.
- `macroRegenerationRefusal` returns an `OfficeExchangeError` and is thrown at both call sites; the lint errors are gone.
- `replacePart` and `replacePartKnownDifferent` share one private `setReplacement` body; the duplication gate is clean.
- XLSX uses `replacePartKnownDifferent` where it already compared bytes; DOCX deliberately keeps equality-checking `replacePart` because it has no prior comparison, so a no-op DOCX save still returns the source bytes.
- `renderAlignedWorksheetRow` rewrites a paired cell's `r` when either the row or column changes, and the structural guard now also covers `autoFilter` and Excel table ranges.
- Deliberate residual limitations: re-anchoring refuses an unequal-count gap when no digest anchor identifies the match; a malformed worksheet whose table relationship or `ref` cannot be resolved contributes no range and does not trigger the shift guard.

## Desktop manual-test fixes (2026-11-20)

The first desktop appointment surfaced five integration defects that the automated suites did not cover; all are fixed and verified live.

- **Save denied outside the Host cwd** — `OfficeDocumentRemote` wrote through `fs` without the per-call `sandboxPolicy`, so the fallback cwd denied the workspace as `FS_SANDBOX_DENIED` and surfaced `conversion-failed`. Save, draft write, and draft delete now resolve the session policy and pass it; a real `fs-sandbox` regression test covers it.
- **Panel not filling the tab body** — a UTF-8 BOM made the bundled first selector `\uFEFF.panel`, dropping `display:flex; flex:1; height:100%`. BOMs were removed from the five affected files; panel and tab body now measure 734 px each.
- **Hidden tabs mounted at 0×0** — Univer was created while its tab cell was `display:none`, producing a blank canvas and no caret. Mounting is now gated on a positive `ResizeObserver` size (environments without `ResizeObserver` keep the old behavior), with tests.
- **Delete but no insert after a tab switch** — DockKit blurred the retained Univer input on hide; on return the browser's editing box had collapsed to height 0, so Chromium disabled `insertText` while command-based Delete still worked. Focus and a caret are now restored on show, and a scoped `min-height:1px` keeps the hidden input's line box. Live repro: switch xlsx→docx, type without clicking; `beforeinput`/`input` fire and undo returns the baseline.
- **Duplicate file names** — the open-document list now replaces the pane tab strip while an office tab is active (`OfficeHeader` under `sidebar.right.pane.header`); non-office tabs keep the strip, and the panel header no longer repeats the filename.

## Convergence gates (2026-11-19)

**Verdict: BLOCKED — the working tree is missing 2,204 tracked files. None of the five gate results below is feature-attributable, and none is reported as a pass.**

### Tree state before the gates

- `git status --short | Measure-Object -Line` → **2,373** lines: **2,204** deleted (` D`), **100** modified (` M`), **69** untracked (`??`).
- `git status --porcelain | Select-String '^ ?D|^D'` (PowerShell equivalent: `git ls-files -d`) → **2,204 deleted tracked files**, all unstaged: `packages/` 2,107, `vendor/` 64, `native/` 33. There are no deletions under `apps/`, `docs/`, `scripts/`, `snapshots/`, `specs/`, `python/`, or `website/`.
- **102 of the 114 affected package/vendor locations lost every tracked file**; the 12 partial locations are tabulated below.
- All 2,204 paths exist in HEAD `f8662e7720` and in the clean main worktree `H:/ForWork` (`git status --porcelain` → 0 lines). Every one of the 100 modified paths still exists on disk, and **none** of them lies inside a wiped location, so the deletion set is disjoint from the feature diff.
- The emptied directories carry last-write times of 2026-11-19 15:55:42–15:55:46 (+08:00), immediately before this run. The worktree reflog has no operation at that time (last entry 2026-11-18 15:08:56), and the deleted count stayed at 2,204 across repeated checks; the damage is external to this run and not a git operation.
- Full path list: `git ls-files -d` (2,204 lines; copy at `D:\kilo\temp\t065\deleted-files.txt`).

Affected locations (114):

```text
native/system
packages/attachment/attachment
packages/boot/app-boot
packages/boot/cmdline
packages/boot/config-editor
packages/boot/hmr
packages/compaction/compaction
packages/compaction/compaction-basic
packages/compaction/compaction-image-offload
packages/context/agent-instructions
packages/core/agent
packages/core/agent-default-model
packages/core/agent-loop
packages/core/scope
packages/core/session
packages/core/system-prompt
packages/core/tool-executor-native
packages/core/tool-executor-ptc
packages/core/tools
packages/credentials/authorization
packages/credentials/credentials
packages/credentials/credentials-local
packages/experimental/hook-protocol
packages/experimental/hooks-claude-code
packages/experimental/hooks-codex
packages/experimental/session-title-all-prompts-llm
packages/fs/fs
packages/fs/fs-local
packages/fs/fs-observation-policy
packages/fs/fs-patch
packages/fs/fs-sandbox
packages/fs/tool-fs
packages/identity/anonymous-user-id
packages/interaction/commands
packages/interaction/permission-presets
packages/interaction/user-approval
packages/interaction/user-questions
packages/jobs/jobs
packages/jobs/jobs-local
packages/jobs/tool-jobs
packages/llm/deepseek-llm-api-extensions
packages/llm/llm
packages/llm/llm-deepseek
packages/llm/llm-deepseek-api-key
packages/llm/llm-pi-ai
packages/llm/plugin-package-inventory-deepseek
packages/llm/token-meter
packages/preset/agent-preset
packages/preset/agent-preset-registry
packages/ptc-runtime/ptc-runtime
packages/sandbox/sandbox
packages/sandbox/sandbox-local
packages/sandbox/sandbox-policy
packages/sandbox/sandbox-windows-acl
packages/session-query/session-query
packages/session/session-format
packages/session/session-format-catalog
packages/session/session-format-v0-to-v1
packages/session/session-format-v1-to-v2
packages/session/session-format-v2-to-v3
packages/session/session-format-v3-to-v4
packages/session/session-persistence
packages/session/session-persistence-jsonl
packages/session/session-projection
packages/session/session-projection-cache
packages/session/session-title
packages/session/session-title-first-prompt-llm
packages/session/session-title-llm
packages/session/tool-working-directory
packages/session/working-directory
packages/settings/settings
packages/shell/bash-local
packages/shell/pwsh-local
packages/shell/shell
packages/shell/shell-env
packages/shell/tool-bash
packages/skill/skill
packages/storage/storage
packages/storage/storage-domain
packages/subagent/subagent
packages/subagent/subagent-spawn-in-process
packages/subagent/tool-subagent
packages/subprocess/subprocess
packages/subprocess/subprocess-local
packages/subprocess/win32-process
packages/test-support/agent-loop-testkit
packages/test-support/llm-replay
packages/test-support/loader-smoke
packages/test-support/tool-testkit
packages/todo/tool-todo
packages/typert/protocol
packages/typert/registry
packages/util/atomic-write
packages/util/brand
packages/util/chunked-list
packages/util/code-language
packages/util/crypto
packages/util/home-paths
packages/util/http-proxy
packages/util/launch-environment
packages/util/lazy-require
packages/util/official-bundles
packages/util/package-manifest
packages/util/time
packages/util/timeout
packages/util/values
packages/workspace/workspace
vendor/cordis
vendor/cosmokit
vendor/group
vendor/include
vendor/loader
vendor/schemastery
vendor/timer
```

Partial locations:

| Location | Deleted / tracked in HEAD |
| --- | ---: |
| `native/system` | 33 / 76 |
| `packages/compaction/compaction-basic` | 18 / 23 |
| `packages/credentials/credentials-local` | 1 / 20 |
| `packages/identity/anonymous-user-id` | 1 / 7 |
| `packages/interaction/user-questions` | 14 / 20 |
| `packages/llm/token-meter` | 28 / 31 |
| `packages/preset/agent-preset` | 16 / 48 |
| `packages/sandbox/sandbox-windows-acl` | 36 / 37 |
| `packages/session/session-projection-cache` | 20 / 22 |
| `packages/skill/skill` | 4 / 15 |
| `packages/subagent/subagent-spawn-in-process` | 15 / 17 |
| `packages/util/official-bundles` | 1 / 10 |

### Gate results

Environment: Windows 11, Node v24.15.0, pnpm 9.15.9, worktree HEAD `f8662e7720`.

| # | Command | Exit | Summary | Classification |
| ---: | --- | ---: | --- | --- |
| 1 | `pnpm run build:artifacts` | 1 | 1 passed, 1 failed, 6 skipped (23.84s). `native addon` passed; `Host Typert` failed (`tsx scripts/generate-typert.ts`) with `TypertAnalysisError: the host program cannot resolve workspace imports` for `@forwork/cordis`, `@forwork/typert-protocol`, `@forwork/cmdline`, `@forwork/schemastery`, and `@forwork/timeout`; host/client compilers, host/client bundles, Desktop bundle, and Web build were skipped. The generator reported no artifact file changed. | Environmental — tree damage; not feature-caused |
| 2 | `pnpm run build:check` | 1 | 0 passed, 2 failed (48.14s). `tsc -p tsconfig.host.json` and `tsc -p tsconfig.client.json` exit 2 with `TS6053 File ... not found` for absent project directories (`vendor/cordis`, `vendor/cosmokit`, `vendor/schemastery`, `vendor/loader`, `vendor/include`, `vendor/timer`, `packages/typert/protocol`, `packages/core/{agent,session,tools}`, `packages/jobs/jobs`, `packages/util/{crypto,timeout}`, and others). | Environmental — tree damage |
| 3 | `pnpm run hygiene` | 1 | 13 passed, 13 failed (46.74s). Failed: `vendor rescope`, `constraints`, `default product isolation`, `product package classification`, `Official bundle catalog`, `pi-ai provider metadata`, `package dependencies`, `optional dependency imports`, `client packages`, `client route resolution`, `plugin record callers`, `no new unknown casts`, `Cordis config`. Every sampled failure is a missing deleted module. | Environmental — tree damage |
| 4 | `pnpm run typecheck` | 1 | 0 passed, 1 failed, 4 skipped (8.29s). Same `Host Typert` missing-import failure, so all four compiler/aggregate checks were skipped. | Environmental — tree damage |
| 5 | `pnpm run duplication` | 0 | `Found 0 clones`; 429 tsx + 2,153 ts = 2,582 files, 515,736 lines, 3,046,198 tokens. | Exit 0 but **not valid evidence**: the sweep saw 468 fewer source files than the 3,050-file run recorded above because the wiped packages are absent. |

Failure evidence beyond the first table rows:

- `build:artifacts` Typert detail: `@forwork/cordis` is imported by `packages/api/account-controller/src/index.ts`, and `@forwork/typert-protocol` by `packages/api/gateway/src/access.ts`; `packages/api/gateway/src/index.ts` itself is clean (`git status --porcelain` empty) and intact — the resolved packages are gone.
- `hygiene` representative errors: `ERR_MODULE_NOT_FOUND` for `packages/llm/llm-pi-ai/scripts/gen-provider-data.ts` (`pi-ai provider metadata`), `packages/boot/app-boot/src/profile.ts` (`constraints`), `packages/boot/app-boot/src/package-meta.ts` (`Official bundle catalog`), `vendor/loader/src/config/diff.ts` (`package dependencies`), and `@forwork/cordis-plugin-include` (`product package classification`, `default product isolation`). Typert additionally reports a `TS2307` cascade such as `packages/experimental/action-cards/src/index.ts: Cannot find module '@forwork/cordis'`.
- No command in this run wrote source: the Typert generator reported no artifact change, and the remaining gates only read the tree.

Classification summary:

- **Feature-caused: none.** Every observed failure names a deleted path or a workspace package resolved through one, and 0 of the 100 modified feature paths lies inside a wiped location. This does not clear the feature; it means these gates never exercised it.
- **Pre-existing: not determinable from this run.** The gates stop before reaching checks that could separate pre-existing gaps, so no pre-existing classification is recorded here.
- **Environmental: all five rows.** The worktree is missing its committed dependency graph, so none of the five commands measures the feature diff.

### Stale referenced-project note

`pnpm exec tsc -b` → exit 1: eight `error TS5083: Cannot read file '<worktree>/tsconfig.json'` for `vendor/cordis`, `vendor/schemastery`, `packages/typert/protocol`, `packages/llm/llm`, `packages/core/{tools,agent,session}`, and `packages/session/session-projection`, followed by a TypeScript 6.0.3 internal `TypeError: Cannot read properties of undefined (reading 'path')` in `hasSameBuildInfo`.

The fresh `build:artifacts` does **not** resolve this. It stops at `Host Typert` before any `tsc -b` stage, and the referenced project directories are absent rather than stale; this is missing input, not stale build info. A clean `build:artifacts` cannot be validated until the deleted files are restored.

### Remediation

- Restore only the deleted tracked paths from HEAD, for example `git ls-files -d | git checkout-index --stdin -f`. That restores committed bytes and leaves the 100 modified and 69 untracked feature paths untouched. It was not run here because this task forbids source edits.
- Then rerun in order: `build:artifacts`, `build:check`, `hygiene`, `typecheck`, `duplication`. Rerun `duplication` after restoration because its current file count is reduced.

### Restored tree rerun (2026-11-19)

The 2,204 deleted tracked paths were restored and the five gates were rerun. The earlier damage record above is unchanged. The restore recovered committed bytes only; it did not recover uncommitted edits carried by deleted files (see "Lost uncommitted edit" below).

**Restore**

- Before: `git status --short` → **2,373** lines (**2,204** ` D`, **100** ` M`, **69** `??`). `git ls-files -d` was saved to `D:\kilo\temp\t065\deleted-before-restore.txt` (**2,204** paths) before restoring.
- Command: `git ls-files -d | git checkout-index --stdin -f` → exit **0**.
- After: `git status --short` → **169** lines (**100** ` M`, **69** `??`), and `git ls-files -d` → **0** lines. All 169 remaining entries are feature paths: the first 120 lines were inspected and the rest are `packages/document/office-exchange/tests/*` plus `specs/002-in-place-ooxml-save/`. No feature path is deleted.
- Dependencies were damaged as well: this worktree's `node_modules` is a junction to `H:\ForWork\node_modules`, whose root links were missing. `corepack pnpm install --frozen-lockfile` (pinned `pnpm@11.28.5`, exit **0**) restored them. Bare `pnpm` 9.15.9 cannot install this workspace: it exits with `ERR_PNPM_LOCKFILE_CONFIG_MISMATCH` because it ignores the workspace `overrides`. `pnpm-lock.yaml`, `pnpm-workspace.yaml`, and `package.json` are unmodified.

**Gate results**

Environment: Windows, Node v24.15.0, pinned pnpm 11.28.5 via `corepack pnpm`, worktree HEAD `f8662e7720`.

| # | Command | Exit | Summary | Classification |
| ---: | --- | ---: | --- | --- |
| 1 | `pnpm run build:artifacts` | 1 | 2 passed, 1 failed, 5 skipped (20.75s). `native addon` and `Host Typert` passed; `host compiler` exited 2 with four errors: `DocumentSaveDisposition` is not exported by `@forwork/session` (`packages/session/session-document/src/document-facts.ts:20`, `packages/document/office-exchange/src/remote.ts:20`), and `'"docm"' is not assignable to type 'DocumentFormat'` (`remote.ts:144`, `remote.ts:212`). Host/client bundles, Desktop bundle, client compiler, and Web build were skipped. | Feature — lost edit (see below) |
| 2 | `pnpm run build:check` | 1 | 1 passed, 1 failed (94.70s). `client aggregate check` passed; `host aggregate check` exited 2 with four errors: `DocumentSaveDisposition` missing in `packages/session/session-document/tests/document-facts.spec.ts:4`; `disposition` missing on `DocumentSaveEventData` in `packages/document/office-exchange/tests/remote.spec.ts:696`; `'"docm"'` and `'"xlsm"'` not assignable to `DocumentFormat` in `packages/session/session-document/tests/document-projection.spec.ts:303,309`. | Feature — lost edit |
| 3 | `pnpm run hygiene` | 1 | 24 passed, 2 failed (156.30s). `package dependencies` exited 1: `TypertAnalysisError` from the same four `office-exchange/src/remote.ts` errors. `publint` exited 1 on `lib/index.js` and other build outputs that do not exist because `build:artifacts` failed; its `pkg.exports["./src/*"] ... does not match any files` messages are warnings (`EXPORTS_GLOB_NO_MATCHED_FILES`), not gate failures. | Feature (`package dependencies`) + environmental (`publint`, missing build output) |
| 4 | `pnpm run typecheck` | 1 | 1 passed, 1 failed, 3 skipped (41.51s). `host compiler` exited 2 with the same four errors as row 1; the client compiler and both aggregate checks were skipped. | Feature — lost edit |
| 5 | `pnpm run duplication` | 0 | `Found 0 clones`; 3,050 files (429 tsx + 2,621 ts), 623,379 lines, 3,573,315 tokens. | Pass |

**Lost uncommitted edit**

The restore cannot recover feature edits that lived in files the wipe deleted. `packages/core/session/src/document-events.ts` is one: it is line 336 of `D:\kilo\temp\t065\deleted-before-restore.txt`, so `checkout-index` rewrote it with HEAD bytes (its last-write time matches the restore), and `git status --porcelain packages/core/session/src/document-events.ts` is now empty. HEAD does not define `DocumentSaveDisposition` (`DocumentFormat` is still `'docx' | 'xlsx' | 'pptx'`), yet the intact modified `docs/persistence-schema.json` names `packages/core/session/src/document-events.ts#DocumentSaveDisposition` at line 225, and the feature contract `contracts/document-save-disposition.contract.md` specifies both the interface and the optional `DocumentSaveEventData.disposition` field. Every remaining compile error traces to this absence.

Recovery attempts found no copy: `git stash list` is empty; the 155 unreachable blobs reported by `git fsck --unreachable --no-reflogs` contain no occurrence of `DocumentSaveDisposition`; the sibling worktrees (`H:/ForWork`, `frame-2`, `frame-workbench-shell`, `logo`, `shimmer-guest`) all carry the same HEAD bytes without the symbol.

Classification summary:

- **Feature (wipe-derived):** rows 1, 2, and 4, and the `package dependencies` part of row 3. All four errors are in feature-modified files (`git status --porcelain packages/document/office-exchange/src/remote.ts packages/session/session-document/src/document-facts.ts` → ` M`), and all depend on the lost edit in the restored `packages/core/session/src/document-events.ts`.
- **Environmental:** `publint` inside row 3 — its errors are missing `lib/` outputs from the failed build, not source defects.
- **Pre-existing:** none newly identified; `duplication` passes and the earlier gates now reach the feature.
- No modified or untracked feature file was overwritten by the restore; the install touched only `node_modules`.

### Final T065 run (2026-11-19)

**Restored session event types verified.** `packages/core/session/src/document-events.ts` (line 42) defines `DocumentFormat = 'docx' | 'docm' | 'xlsx' | 'xlsm' | 'pptx'`, exports `DocumentSaveDisposition` (line 218), and adds the optional `disposition?: DocumentSaveDisposition` field to `DocumentSaveEventData` (line 252). The barrel at `packages/core/session/src/index.ts:29` re-exports `./document-events.ts`. `git status --porcelain packages/core/session/src/document-events.ts` → ` M`; this is the intended restored feature edit, and every consumer compiles against it.

Environment: Windows, Node v24.15.0, pnpm 9.15.9 (commands run exactly as listed), worktree HEAD `f8662e7720`. Before the run, `git status --short` → 170 lines (101 ` M`, 69 `??`) and no tracked deletions; the 101st modified file is the restored `document-events.ts` relative to the previous 100-file state.

| # | Command | Exit | Summary | Classification |
| ---: | --- | ---: | --- | --- |
| 1 | `pnpm run typecheck` | 0 | 5 passed, 0 failed, 0 skipped (128.69s): `Host Typert`, `host compiler`, `client compiler`, `host aggregate check`, `client aggregate check`. | Pass |
| 2 | `pnpm exec vitest run packages/session/session-document/tests packages/document/office-exchange/tests` | 0 | 55 test files passed, 1,210 tests passed (39.41s). | Pass |
| 3 | `pnpm run build:artifacts` | 1 | 4 passed, 1 failed, 3 skipped (44.46s). `native addon`, `Host Typert`, `host compiler`, and `Host bundle` passed; `Desktop bundle` failed because `vendor/schemastery/lib/index.mjs` imports `@forwork/cosmokit`, which rolldown could not resolve; `client compiler`, `Client bundle`, and `Web build` were skipped as dependents. | Environmental — dangling workspace links in shared `node_modules` |
| 4 | `pnpm run build:check` | 0 | 2 passed, 0 failed (56.21s): `host aggregate check`, `client aggregate check`. | Pass |
| 5 | `pnpm run hygiene` | 1 | 25 passed, 1 failed (259.83s). Only `publint` failed: `packages/typert/registry` has five missing-file errors for `lib/index.js`, `lib/client.node.js`, and `lib/client.browser.js`; `package dependencies`, `vendor rescope`, `no new unknown casts`, and the other 22 gates passed. | Environmental — missing output of the skipped `Client bundle` |
| 6 | `pnpm run duplication` | 0 | `Found 0 clones`: 3,050 files (429 tsx + 2,621 ts), 623,400 lines, 3,573,388 tokens (0.67s). | Pass |

**Row 3 evidence.** `vendor/schemastery/node_modules` and `vendor/cordis/node_modules` are junctions into the sibling main worktree (`H:\ForWork\vendor\...`), but the `@forwork/*` links inside them are relative targets containing `.kilo\worktrees\protective-talk`. Resolved from this worktree's logical path they become `H:\ForWork\.kilo\worktrees\protective-talk\.kilo\worktrees\protective-talk\...` and do not exist. Four links are dangling (cosmokit under `schemastery` and `cordis`; `cordis-plugin-include`, `cordis-plugin-loader`). Node confirms `ERR_MODULE_NOT_FOUND: Cannot find package '@forwork/cosmokit' imported from ...\vendor\schemastery`. `git status --porcelain vendor/schemastery vendor/cosmokit apps/desktop apps/desktop/scripts` is empty: no feature path participates. Rolldown therefore leaves the bare import external, and the `desktop-packaged-imports` guard fails the bundle by design.

**Row 5 evidence.** `packages/typert/registry/tsdown.config.ts` calls `clientBundle(..., { dualClient: true })`, which emits `index.js`, `client.node.js`, and `client.browser.js` in the `Client bundle` gate that row 3 skipped. `git status --porcelain packages/typert/registry` is empty. The `pkg.exports["./src/*"] ... does not match any files` messages throughout the log are publint warnings; the five missing-file messages above are the errors that set exit 1.

Classification summary:

- **Feature-caused: none observed.** The restored event types compile and both focused suites pass; no failure names a feature path, and none of the 101 modified or 69 untracked feature paths lies in the failing modules.
- **Environmental:** row 3 (dangling links in the shared worktree `node_modules`) and row 5 (missing `Client bundle` output). Both failures originate in unmodified vendor/build paths, and neither error names feature code.
- **Pre-existing:** not separately measurable here, because row 3 fails on worktree link state before Desktop packaging and row 5 depends on row 3's skipped Client bundle.

Final state: 4 of 6 commands exit 0 (typecheck, focused vitest, build:check, duplication); `build:artifacts` and `hygiene` exit 1 for the environmental reasons above. This run edited no source; logs are in `D:\kilo\temp\t065\`. No pass is claimed for exit != 0.

### T064 partitioned coverage (2026-11-19)

Command (exact):

    $env:DSH_COVERAGE_PARTITIONS='4'; pnpm run test:coverage:partitioned; $code=$LASTEXITCODE; Remove-Item Env:DSH_COVERAGE_PARTITIONS; exit $code

**Result: not completed; no pass claimed.** The process tree was killed when the 30-minute tool budget expired. The harness returned `Unknown: ChildProcess.kill`; the shell never reached `$LASTEXITCODE`, so **no exit code exists** (it must not be read as 0). Wall time to the kill: 18:15:29 to 18:46:30 (+08:00), about 31m01s. Log: `D:\kilo\temp\t064\run.log` (635,779 bytes, 7,734 lines).

**Merged aggregate: not produced.** `coverage/.partitioned/blobs/` has 0 entries, no `partition-*.times.json` was written, and the coordinator never reached `assertCompleteBlobSet` or the `--merge-reports` command. Statements/branches/functions/lines and the per-file 100% list are therefore unavailable. No file below 100% can be named from this run.

**Partitions:** `coverage-partitions: start partition {1,2,3,4}/4` printed for all four. No partition reported a startup failure, no `coverage-partitions: FAIL` line, and no partition-level timeout message appeared. The only blocker is the overall wall-clock budget. The blocking "error" is the harness kill, not a runner message: there is no 30-line runner error to quote, and the final 30 log lines are normal test progress (reproduced below).

Progress at the kill: inventory 2,634 test files; 805 test files passed and 34 failed (121 failed cases), so 839 files (31.9%) reported. For the feature packages, 32 `office-exchange`, 6 `ui-office-editor`, and 13 `core/session` test-file results appeared, all passing; 0 `session-document` results appeared before the kill. Coverage for those packages is unavailable because the merge never ran; this is not evidence that they are at 100%.

All 34 failed test paths return an empty `git status --porcelain <path>`. The feature roots themselves are modified: `packages/document/office-exchange` 84 entries, `packages/client/ui-office-editor` 24, `packages/session/session-document` 5, `packages/core/session` 1; no failed test path is under any of them. Classifications below combine a focused non-coverage rerun (six files; exit 1; 12 failed / 611 passed / 6 skipped; 73.49s; `D:\kilo\temp\t064\focused.log`) with the paths and durations in the killed run.

| Failed test path (grouped) | Files/cases | Classification and evidence |
| --- | --- | --- |
| `packages/core/tools/tests/gen-tool-catalog.spec.ts` | 1/1 | **Pre-existing (diagnosed).** Received 75 tools vs 71 expected; the four extras are `document_convert`, `document_edit_cells`, `document_edit_paragraph`, `document_read`, registered in HEAD's `packages/document/tool-office/src/index.ts` (`git grep HEAD`), while the HEAD spec omits them (`git show HEAD:...` has no `document_`). Spec and `src/index.ts` are clean (`git status --porcelain` empty). |
| `scripts/run-gates.spec.ts` | 1/2 | **Pre-existing (diagnosed) + environmental (diagnosed).** `brand-assets` is present in HEAD `package.json` and `scripts/brand-assets.manifest.json` (commit 78e6a26f15) but absent from the clean spec's expected list. The second case, "kills the child when the abort signal fires", timed out at the 5s test limit after 71.6s. Spec and gate sources clean. |
| `scripts/verify-default-product-isolation.spec.ts` | 1/6 | **Pre-existing (diagnosed).** Six assertions expect `unknown workspace package` failures that the verifier does not report; spec and `scripts/verify-default-product-isolation.ts` are both clean (`git status --porcelain` empty). |
| `scripts/verify-package-dependencies.spec.ts` | 1/2 | **Pre-existing (diagnosed).** Export classification omits `@forwork/lazy-require` (`*`, `createLazyRequire`) and diffs the rest; spec and `scripts/verify-package-dependencies.ts` clean. |
| `scripts/verify-package-meta.spec.ts` | 1/1 | **Pre-existing (diagnosed).** SVG convention check reports two problems for the fixture; spec and `scripts/verify-package-meta.ts` clean. |
| `scripts/locale-dictionary-parity.spec.ts` | 1/1 | **Environmental (diagnosed).** `Error: Test timed out in 5000ms` after 12.0s (18.2s in the coverage run); the dictionary assertion never ran. Spec clean. |
| `packages/subagent/subagent-acp/tests/subagent-acp.spec.ts` | 1/11 | **Environmental (inferred from case names and durations).** EOF/TERM process-tree cases waited 61.8s and 123.4s; file total 896.0s. Windows process-signal semantics; path clean. |
| `packages/experimental/client-ui-forge/tests/controller.host.spec.ts`, `worktree.host.spec.ts` | 2/30 | **Environmental (inferred).** Real Git/worktree host suite, 313.3s and 60.3s; paths clean. |
| `packages/experimental/worktree/tests/worktree.spec.ts` | 1/7 | **Environmental (inferred).** Real Git worktree suite, 133.5s; path clean. |
| `packages/boot/plugin-market/tests/repository-cache.spec.ts` | 1/8 | **Environmental (inferred).** Real Git/pnpm repository cache, 91.6s; path clean. |
| `packages/boot/plugin-manager/tests/manager.spec.ts` | 1/3 | **Environmental (inferred).** pnpm link/manager host suite, 75.8s; path clean. |
| `packages/boot/app-boot/tests/profile-resolution.spec.ts` | 1/12 | **Environmental (inferred).** 88.3s path-resolution suite; path clean. |
| `packages/boot/app-boot/tests/plugin-compatibility.spec.ts` | 1/2 | **Pre-existing (inferred).** Fast version/plugin assertions in a clean spec; no feature symbol. |
| `packages/boot/app-boot/tests/app-boot.spec.ts` | 1/1 | **Undetermined.** Clean path; failure detail was lost with the kill. |
| `packages/bundle/headless/tests/headless.spec.ts` | 1/1 | **Pre-existing (inferred).** Fast bundle assertion in a clean spec. |
| `packages/credentials/deepseek-account-platform/tests/account.spec.ts` | 1/1 | **Environmental (inferred).** Real endpoint/account suite, 9.8s; path clean. |
| `packages/deliverables/workspace-changes/tests/plugin.spec.ts` | 1/2 | **Environmental (inferred).** Git changes host suite, 37.3s; path clean. |
| `packages/experimental/liveshare/tests/shareable-catalog.spec.ts`, `viewer-brand-registration.client.spec.tsx`, `viewer-shell-artifact.client.spec.tsx` | 3/3 | **Undetermined / environmental (inferred).** Brand/artifact and 49.6s shell-artifact cases; paths clean. |
| `packages/experimental/translator/tests/deepseek.spec.ts` | 1/2 | **Environmental (inferred).** Real translator endpoint; path clean. |
| `packages/llm/llm-deepseek/tests/adapter.spec.ts`, `loader-composition.spec.ts` | 2/4 | **Environmental (inferred).** Real model/credential routes; paths clean. |
| `packages/llm/llm-pi-ai/tests/adapter.spec.ts` | 1/1 | **Environmental (inferred).** Real provider route; path clean. |
| `packages/ptc-runtime/ptc-runtime-node/tests/runtime.spec.ts` | 1/3 | **Undetermined.** Clean path; sandbox/process failure detail lost. |
| `packages/sdk/client/tests/sdk-client.spec.ts` | 1/2 | **Undetermined.** Clean path; wire/process failure detail lost. |
| `packages/session/session-telemetry-otel/tests/otel.spec.ts` | 1/1 | **Undetermined.** "Real SDK pipeline" assertion in a clean path; detail lost. |
| `packages/subagent/subagent-dsh-sdk/tests/subagent-dsh-sdk.spec.ts` | 1/1 | **Undetermined.** Clean path; detail lost. |
| `packages/workflow/workflow-ptc/tests/workflow-ptc.spec.ts` | 1/1 | **Environmental (inferred).** PTC sandbox host suite, 17.5s; path clean. |
| `scripts/local-bundle-registry.spec.ts` | 1/2 | **Environmental (inferred).** Real pnpm install/store; path clean. |
| `scripts/release/pack.spec.ts` | 1/7 | **Environmental (inferred).** Real pnpm/Git packaging, 94.7s; path clean. |
| `scripts/windows-package-workflow.spec.ts` | 1/2 | **Environmental (inferred).** Windows packaging workflow, 20.3s; path clean. |
| `apps/desktop/tests/cua-optional-package.spec.ts` | 1/1 | **Environmental (inferred).** Downloads the macOS optional tarball; path clean. |

**Feature classification:** none proven. The only feature-adjacent failure, `gen-tool-catalog`, is expectation drift already present at HEAD (the working-tree edits to `edit.ts`, `edit-cells.ts`, and `read.ts` do not change the registered tool names). No failing test path lies under the modified feature roots.

The focused rerun also shows `scripts/locale-dictionary-parity.spec.ts` and the `run-gates` abort case failing only on the 5s test timeout, so those must not be read as dictionary or gate-logic regressions.

Blocking error (verbatim, the only error text):

    Unknown: ChildProcess.kill

No runner error was printed; the first 30 lines of a runner error do not exist. The final 30 lines of partial output, all normal progress, follow:

```
     ✓ reports an unusable output directory without attempting an ACL write  4206ms
     ✓ reports an unreadable ACL as a partial observation without inventing permissions  4805ms
     ✓ treats missing observations as incomplete and refuses before changing anything  4610ms
     ✓ refuses to restore a record for another path or a path outside -AllowRoot  11828ms
     ✓ repairs its own workspace root when -AllowRoot is the object itself  6709ms
     ✓ is idempotent: a second repair run changes nothing  6070ms
 ✓ |thread-safe| apps/desktop/tests/windows-timestamp.spec.ts (13 tests) 565ms
 ✓ |thread-safe| packages/workflow/workflow-ptc/tests/realm.spec.ts (15 tests) 39ms
 ✓ |thread-safe| packages/client/ui-open-in-app/tests/session-open-with.client.spec.ts (9 tests) 325ms
 ✓ |thread-safe| packages/client/ui-primitives/tests/tooltip.client.spec.tsx (47 tests) 806ms
 ✓ |thread-safe| packages/api/session-controller/tests/summary-list.host.spec.ts (3 tests) 125ms
 ✓ |thread-safe| packages/experimental/liveshare/tests/source-access-lifetime.spec.ts (1 test) 136ms
 ✓ |thread-safe| packages/boot/hmr/tests/reconciliation.spec.ts (3 tests) 497ms
 ✓ |thread-safe| apps/desktop/tests/desktop-publish.spec.ts (12 tests) 35ms
 ✓ |thread-safe| packages/experimental/liveshare/tests/viewer-settings-privacy.spec.ts (4 tests) 98ms
 ✓ |thread-safe| packages/client/ui-sidebar/tests/sidebar-root.client.spec.tsx (46 tests) 1294ms
 ✓ |thread-safe| packages/experimental/liveshare/tests/viewer-reader-peer.spec.ts (6 tests) 341ms
 ✓ |thread-safe| scripts/persistence-schema-model.spec.ts (8 tests) 18ms
 ✓ |thread-safe| packages/api/session-controller/tests/commands-upload-file.host.spec.ts (37 tests) 676ms
 ✓ |thread-safe| packages/client/ui-sidebar/tests/scroll-handoff.client.spec.ts (16 tests) 28ms
 ✓ |thread-safe| packages/client/ui-theme/tests/client-styles.client.spec.ts (11 tests) 83ms
 ✓ |thread-safe| packages/experimental/forge-session/tests/commands.spec.ts (35 tests) 17ms
 ✓ |thread-safe| packages/client/ui-office-editor/tests/adapter-runtime.client.spec.ts (5 tests) 27ms
 ✓ |thread-safe| packages/client/ui-conversation/tests/queue-dock.client.spec.tsx (36 tests) 1096ms
 ✓ |thread-safe| packages/util/workspace-path/tests/file-address.spec.ts (48 tests) 11ms
 ✓ |thread-safe| apps/desktop/tests/preload-menu.client.spec.ts (8 tests) 341ms
 ✓ |thread-safe| packages/workflow/tool-workflow/tests/tool-workflow.spec.ts (33 tests) 3238ms
       ✓ an abort releases the tool even when the script parks on a promise no hook owns  1071ms
       ✓ a background run over the sandboxed PTC engine settles its job with the rendered return value  572ms
 ✓ |thread-safe| packages/client/ui-message-feedback/tests/feedback-dialog.client.spec.tsx (10 tests) 388ms
```

No source file was edited. The only writes were the gitignored `coverage/` tree created by the runner, logs under `D:\kilo\temp\t064\`, and this appended section.

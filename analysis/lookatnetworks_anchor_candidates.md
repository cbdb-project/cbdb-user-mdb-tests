# LookAtNetworks anchor candidate inventory (PR AQ)

Pure-pyodbc inventory of candidate `c_personid` anchors for a future LookAtNetworks CmdRun fixture experiment.  The current matrix uses Zhu Xi (2 471 associations) which times out; this list ranks smaller candidates by likely 1-hop expansion cost.

Inputs: ASSOC_DATA + KIN_DATA + BIOG_MAIN, User MDB at `data\CBDB_BJ_User.mdb`.  No Access COM.  No Networks-form static knowledge of the actual depth/loop caps was used — these are conservative estimates of 1-hop reach.

## Headline counts

- Candidate persons (assoc-degree 5..50): **4253**
- flag `likely_safe_under_120s`: 1150
- flag `medium`: 1095
- flag `too_large_kin`: 5
- flag `too_large`: 2002
- flag `too_large_known`: 1
- Of all candidates, in current test_inputs.json: 36

## Top recommended anchors (likely_safe_under_120s)

Sorted by: in_test_inputs (preferred), has_name_chn, has_index_year, then closeness to assoc-degree 10.  Pick 1–3 from this table for the next Access COM experiment.

| # | c_personid | name (chn) | name (py) | assocs | kin | est 1-hop | dyn | index_year | in_test_inputs |
|---:|---:|---|---|---:|---:|---:|---:|---:|---|
| 1 | 4 | 查道 | Zha Dao | 5 | 9 | 99 | 15 | 955 | yes |
| 2 | 1368 | 明鎬 | Ming Hao | 5 | 0 | 110 | 15 | 989 | yes |
| 3 | 3135 | 張君平 | Zhang Junping | 5 | 2 | 31 | 15 | 964 | yes |
| 4 | 29794 | 焦循 | Jiao Xun | 17 | 5 | 165 | 20 | 1763 | yes |
| 5 | 828 | 胡旦 | Hu Dan | 10 | 4 | 64 | 15 | 948 | no |
| 6 | 1882 | 王隨 | Wang Sui | 10 | 1 | 99 | 15 | 972 | no |
| 7 | 30270 | 曹植 | Cao Zhi (2) | 10 | 1 | 10 | 26 | 192 | no |
| 8 | 31628 | 岑文本 | Cen Wenben | 10 | 6 | 166 | 6 | 595 | no |
| 9 | 31957 | 崔湜 | Cui Shi | 10 | 4 | 182 | 6 | 671 | no |
| 10 | 33917 | 彭定求 | Peng Dingqiu | 10 | 5 | 58 | 20 | 1645 | no |

## Reference: known-large anchors

| c_personid | name (chn) | assocs | kin | est 1-hop | flag |
|---:|---|---:|---:|---:|---|
| 147 | 錢公輔 | 11 | 2 | 1453 | `too_large` |
| 560 | 范純禮 | 11 | 7 | 941 | `too_large` |
| 1464 | 施昌言 | 9 | 7 | 827 | `too_large` |
| 146 | 錢易 | 8 | 11 | 1136 | `too_large` |
| 125000 | 方逢時 | 12 | 8 | 507 | `too_large` |
| 1 | 安惇 | 13 | 5 | 814 | `too_large` |
| 62 | 陳康伯 | 13 | 7 | 1090 | `too_large` |
| 1504 | 蘇師德 | 7 | 11 | 1594 | `too_large` |
| 1897 | 王存 | 7 | 8 | 1026 | `too_large` |
| 118 | 程琳 | 6 | 14 | 1065 | `too_large` |

## Caveats

- ASSOC_DATA edges are deduplicated bidirectionally by the script; the actual LookAtNetworks CmdRun may walk them differently (with kin / depth filters).  These counts are upper bounds for the 1-hop reach.
- `est_1_hop_assoc_total` sums the degrees of every 1-hop neighbor — this is a worst-case for a depth-2 walk that doesn't dedupe second-hop neighbors.  CmdRun does dedupe (per `Form_LookAtNetworks.vb`'s ZZ_SCRATCH_PEOPLE write pattern), so true cost will be lower.
- We did NOT filter by `gMaxFilterTotal=29` etc. (the default checkbox state from Form_Open).  The default Networks UI applies association-type filters that would shrink the candidate set further.
- The 1-hop estimator assumes uniform degree; high-variance neighbours (e.g. one neighbour is a Zhu-Xi-scale hub) can blow past `est_1_hop_assoc_total`.  When picking from the table, consider whether the anchor's named neighbors are likely hubs.

## How to use the recommended list

Pick 1–3 anchors from the recommended table.  For each, the next Access COM experiment should:

1. Open LookAtNetworks (Form_Open is fine per PR AA).
2. Set the picker to that c_personid via `set_picker_codes("ZZ_SCRATCH_IMPORT_PEOPLE", [pid])`.
3. Set `gMaxNodes` and `gMaxLoops` to small values (e.g. 20 nodes / 1 loop) before firing CmdRun.
4. Use `Form_Timer` trigger with a 120 s budget.
5. If CmdRun completes, capture the resulting ZZ_SCRATCH_PEOPLE / ZZ_SOCIAL_NETWORK row counts.

Out-of-scope for this PR.  See `analysis/lookatnetworks_form_open_hang.md` for what we already know about the form's behaviour under COM.
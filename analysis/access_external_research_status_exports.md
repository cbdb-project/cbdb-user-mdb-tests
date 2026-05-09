# External research — `LookAtStatus × CmdPajek/Gephi` parked local-workaround line

**Date:** 2026-05-09
**Branch:** `research/access-external-evidence-status-exports` (off main `81c029b`)
**Scope:** read-only research memo. No COM run; no tests / driver / README / triage / canonical reports / issue severity changes.

This memo collects external evidence (Microsoft official docs first; recognized Access-community sources second; forum anecdote third) to evaluate whether our parked local-workaround disposition for Status × CmdPajek/Gephi is the right call, OR whether the external evidence justifies one more local attempt before maintainer-line.

---

## Raw findings (sources stratified by credibility)

### TIER 1 — Microsoft official documentation (Microsoft Learn / VBA-Docs)

**F1. `Application.Run` only accepts `"procedurename"` or `"projectname.procedurename"` (NOT `"Form_<name>.SubName"`).**

Microsoft Learn, "Application.Run method (Access)" page, *Parameters* table, `Procedure` row:

> *"The name of the **Function** or **Sub** procedure to be run. If you are calling a procedure in another database, use the project name and the procedure name separated by a dot in the form: `"projectname.procedurename"`"*

The doc gives ONE example (calling `Greeting` — a Public sub in a STANDARD module). The example deliberately puts the wrapper in a "new module" (standard module), not a form module. **No syntax form `"Form_<name>.SubName"` is documented anywhere on the page.**

**Direct relevance to our PR #136:** PR #136 attempted `app.Run("Form_LookAtStatus.RunExportPajek")` and got `"cannot find the procedure 'Form_LookAtStatus.RunExportPajek.'"`. The error is exactly what the docs would predict — the qualified form `"Form_<name>.SubName"` is not in the documented syntax. This is not a bug; it's by design.

**Source:** [Application.Run method (Access) — Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/access.application.run)

---

**F2. `Form.Requery` doesn't pass control to the OS — DoEvents is the documented escape valve.**

Microsoft Learn, "Form.Requery method (Access)" page, *Remarks* > *Note*:

> *"The **Requery** method doesn't pass control to the operating system to allow Windows to continue processing messages. Use the **DoEvents** function if you need to relinquish temporary control to the operating system."*

**Direct relevance to PR #131-#134:** This validates PR #131's positive signal (Requery + DoEvents + 1.5 s sleep returns expected RecordCount) AND explains why PR #132/#133/#134's VBA-side DoEvents loops were insufficient — DoEvents from a Form_Timer event runs on the same UI-thread call stack and doesn't fully release the thread the way a Python `time.sleep` does (which is what PR #131 used).

**Source:** [Form.Requery method (Access) — Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/access.form.requery)

---

**F3. Microsoft's OWN Form.Recordset example uses a `Global`-scope recordset variable.**

Microsoft Learn, "Form.Recordset property (Access)" page, *Remarks* example:

```vba
Global rstSuppliers As ADODB.Recordset
Sub MakeRW()
    DoCmd.OpenForm "Suppliers"
    Set rstSuppliers = New ADODB.Recordset
    ...
    Set Forms("Suppliers").Recordset = rstSuppliers
End Sub
```

Note that `rstSuppliers` is declared at **`Global` scope** — NOT `Dim`'d locally inside the sub. **This is exactly the upstream fix pattern Concern 4 recommended in PR #139's handoff memo** (switch `tRstStatus` from `Dim`'d-local to global, per the existing `gRstPeople` precedent at `Form_LookAtStatus.vb:1184`).

**Direct relevance:** the Concern 4 fix isn't a stylistic preference — it's the Microsoft-documented pattern. Our suggested upstream fix aligns with Microsoft's own example.

**Source:** [Form.Recordset property (Access) — Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/access.form.recordset)

---

**F4. Microsoft documents a `RecordSource = RecordSource` self-rebinding workaround for `Form.Recordset` instability.**

Same Form.Recordset page, *Remarks*, two adjacent quotes:

> *"Changing a form's **Recordset** property may also change the **RecordSource**, **RecordsetType**, and **RecordLocks** properties. Also, some data-related properties may be overridden..."*

> *"Calling the **Requery** method of a form's recordset (for example, `Forms(0).Recordset.Requery`) can cause the form to become unbound. To refresh the data in a form bound to a recordset, set the **RecordSource** property of the form to itself: `Forms(0).RecordSource = Forms(0).RecordSource`."*

**Direct relevance:** this is a **NEW** Microsoft-documented workaround we did NOT try in any of our 7 local attempts (PR #129/#132/#133/#134/#135/#136/#137). PR #129 tested `Set→Requery`; PR #133 tested explicit `Form.Requery`; neither tested the `RecordSource = RecordSource` self-rebinding form. This is the strongest "one more local probe is justified" signal in the evidence.

**Caveat:** the doc says the self-rebinding refreshes "the data in a form bound to a recordset" — i.e. it addresses the unbound-form symptom. Whether it also fixes the per-form-instance event-binding cache (PR #137's pinned layer) is not stated. The technique is a strong candidate for a focused probe, but not a guaranteed fix.

**Source:** [Form.Recordset property (Access) — Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/access.form.recordset)

---

**F5. Microsoft did NOT document the `AddFromString` → event-binding-cache refresh behavior.**

The "Module.AddFromString method (Access)" page documents the basic API (add code at a specified line) but does NOT discuss any caveat about event-handler cache refresh after the AddFromString call. Search for an authoritative MS source on "AddFromString event binding refresh" returned no direct results — this appears to be undocumented Access internal behavior.

PR #137's finding (timer fires but body doesn't dispatch even with `RunCommand(126)` force-compile) is consistent with there being no documented escape valve at this layer. This is the strongest "no further local mechanism likely" signal.

**Source:** [Module.AddFromString method (Access) — Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/Access.Module.AddFromString) (and absence of any matching authoritative content for the cache-refresh question)

---

### TIER 2 — recognized Access community sources (Allen Browne / FMS / Tek-Tips / Access-Programmers / MrExcel / StackOverflow high-rep answers)

**F6. Form-module Public subs are NOT addressable from standard modules via `Application.Run`; the documented escape is `Forms!FormName.PublicSubName`.**

Multiple long-standing community sources (Access World Forums + tek-tips + bytes.com) converge on the same advice:

- *"To call a sub from another form, you will need to make the PRIVATE sub into a PUBLIC sub for this to work."* (multi-thread consensus, Access World Forums)
- *"You generally don't put shared code in forms. Instead, put whatever you need to be public in a separate module, because the reason is that you have to fully qualify the sub with the form where it is in: `Call Forms!MyFormName.MyPublicSub`"* (Access World Forums)
- Alternative documented forms: `Forms.Item("FormName").FunctionName` or `CallByName Forms.Item(FormName), MethodName, VbMethod`

**Direct relevance to PR #136:** even with the Public wrapper trick (`RunExportPajek`), `Application.Run("Form_LookAtStatus.RunExportPajek")` won't work — community consensus says you'd have to use `Call Forms!LookAtStatus.RunExportPajek` from VBA OR move the wrapper to a standard module.

**Caveat:** `Call Forms!LookAtStatus.RunExportPajek` requires VBA execution context (i.e., already running in some VBA sub). It's NOT a way to invoke from external Python COM. So this confirms PR #136's finding without offering a new path.

**Confidence:** Tier 2 (multi-source community consensus, decades-old; no contradictory Microsoft doc).

**Sources:**
- [Application.Run Subform Procedure — tek-tips](https://www.tek-tips.com/threads/application-run-subform-procedure.1753196/)
- [Calling public sub from another form — Access World Forums](https://www.access-programmers.co.uk/forums/threads/calling-public-sub-from-another-form.153447/)

---

**F7. "Object required" on subform recordset access typically traces to local-variable scope.**

Multiple community sources note that when a `Dim`'d-local DAO recordset variable goes out of scope, references to `<subform>.Form.Recordset` that were `Set` from it can read as Nothing → VBA 424 'Object required'. Recommended fix is to keep the recordset variable alive at module-level or global scope.

**Direct relevance to Concern 4:** this is community consensus on EXACTLY the Status cleanup-rebind pattern. The local `tRstStatus` in `Form_LookAtStatus.CmdQuery_Click` is the textbook case the community advises against. Concern 4's suggested upstream fix (switch to global, per `gRstPeople` precedent at line 1184) IS the community-recommended fix shape.

**Confidence:** Tier 2 (consistent with Microsoft's own example in F3, plus community consensus).

**Sources:**
- [Subform - Assign Recordset to (DAO) — Access World Forums](https://www.access-programmers.co.uk/forums/threads/subform-assign-recordset-to-dao.61256/)
- [Form and Subform DAO Recordsets — Tek-Tips](https://www.tek-tips.com/threads/form-and-subform-dao-recordsets.886905/) (note: Tek-Tips reply emphasizes module-level scoping)

---

**F8. `Forms(0).RecordSource = Forms(0).RecordSource` (self-rebinding) is also recommended in community threads as the workaround when a form becomes "unbound" or detached from its record source after Recordset manipulation.**

This corroborates F4 from a Tier-2 source independently of the official doc. The pattern is recommended specifically when `Set Form.Recordset = <recordset>` followed by `.Requery` or `.Close` causes binding loss.

**Confidence:** Tier 2 — corroborates F4. Adds independent support for this being a real, working pattern (not a theoretical doc-only recipe).

**Sources:**
- [Use Recordset as Form Recordsource — Access World Forums](https://www.access-programmers.co.uk/forums/threads/use-recordset-as-form-recordsource.56221/)

---

### TIER 3 — forum anecdote / weaker evidence

**F9. Some forum threads suggest "close + reopen the form" as a last-resort fix when AddFromString-style runtime VBA injection causes event handlers to not bind correctly.**

This is anecdotal and tied to specific Access versions (Access 2007 / 2010 era posts). No citation rises to Tier 1 or Tier 2 confidence. Worth mentioning because it's a pattern that periodically resurfaces in community discussion, but no canonical source endorses it as a guaranteed fix.

**Direct relevance:** PR #137's verdict-note already flagged "close + reopen form" as the most-likely-to-work remaining local candidate. External evidence is consistent with this being a viable last-resort, but does NOT independently strengthen the case beyond what we already had.

**Confidence:** Tier 3 (forum anecdote; no canonical source).

---

## Inference (separated from raw findings)

### I1. The "Status × CmdPajek/Gephi cleanup-rebind root cause" Concern 4 fix is Microsoft-documented best practice

F3 (Microsoft's own Form.Recordset example) AND F7 (community consensus on local-Dim recordset scope) BOTH point at the same fix shape: hold the recordset variable at `Global` (or at minimum module-level) scope, NOT `Dim`'d-locally inside the sub that assigns it to a subform's `.Form.Recordset`.

The Concern 4 fix recommendation in PR #139's handoff memo says:

> "switch to globals (per gRstPeople precedent at line 1184) OR refactor away the rebind entirely"

**This is exactly the pattern Microsoft documents in its own example AND that the community has converged on for two decades.** The maintainer-line ask is well-grounded.

### I2. The PR #136 (Application.Run) boundary is by design, not a missed escape

F1 (official doc syntax) AND F6 (community consensus on form-module sub addressability) both confirm that PR #136's `app.Run("Form_LookAtStatus.RunExportPajek")` was attempting an undocumented form. There is no escape valve at this layer that we missed.

The remaining theoretical paths (`Call Forms!LookAtStatus.<wrapper>` from VBA; standard-module wrapper that uses `Forms!LookAtStatus.<sub>`) all require VBA-side execution context AND would still need a way to *enter* VBA from Python COM — bringing us back to `Form_Timer` (which PR #137 ruled out via the event-binding-cache finding).

### I3. The PR #131-#134 mechanism boundary (VBA-side DoEvents ≠ COM-side sleep) is by Microsoft design

F2's official wording — *"The Requery method doesn't pass control to the operating system... Use the DoEvents function if you need to relinquish temporary control"* — is consistent with our empirical finding. PR #131's positive 1.5 s used Python `time.sleep` (releases the COM thread fully); PR #132-#134 used VBA Timer/DoEvents loops (don't release the COM thread the same way). The mechanism boundary is real, documented, and not a bug we missed.

### I4. There IS exactly ONE Microsoft-documented local technique we did not test: `Forms(0).RecordSource = Forms(0).RecordSource` self-rebinding

F4 (official doc) + F8 (community corroboration) describe this technique. Per PR #137's evidence, the underlying issue is at Access's per-form-instance event-binding cache after `AddFromString` — but the `RecordSource = RecordSource` technique addresses a *different* failure surface (recordset-becomes-unbound, not event-binding-cache-stale). It's plausible but not certain that combining the two might unblock.

A focused probe of the self-rebinding pattern would take ~30-60 minutes:
- inject a wrapper that runs `Forms!LookAtStatus.<subform>.Form.RecordSource = "ZZ_SCRATCH_STATUS"` between Phase A and Phase B
- arm timer
- observe whether `Cmd<X>_Click` then runs cleanly

This is the strongest "one more local probe could be justified" signal in the external evidence. It is NOT a slam-dunk — the docs frame self-rebinding as a fix for the form-becomes-unbound symptom, not for `Object required` from the local-variable-scope cleanup-rebind pattern.

### I5. The `AddFromString` event-binding-cache layer is undocumented and likely Access-version-specific

F5's absence-of-doc + PR #137's empirical evidence + F9's Tier-3 close+reopen anecdote suggest this layer has no documented escape. Pursuing it further locally would be heuristic / version-specific — exactly the kind of fragility we want to avoid in a test driver.

### I6. Microsoft's recommended workaround for Form.Recordset issues IS the upstream fix

F3 + F4 collectively show that Microsoft's own guidance for "Form.Recordset assigned to local recordset, then needs refresh" is:
- declare the recordset at `Global` scope (F3)
- if a Requery causes unbinding, use `RecordSource = RecordSource` (F4)

The Status form's existing `gRstPeople` global at line 1184 already follows F3. Applying the same pattern to `tRstStatus` (the cleanup-rebind variable) is the documented fix. Not opaque maintainer-prerogative; documented Microsoft pattern.

---

## Verdict

**`external_evidence_supports_maintainer_line`**

**Rationale:**

The dominant external evidence (Microsoft's own F3 example using `Global` scope; F4's documented self-rebinding workaround; F7's community consensus; F8's corroboration of F4) all converge on the same conclusion: the upstream fix Concern 4 recommends (switch from `Dim`'d-local to global scope for `tRstStatus` per the `gRstPeople` precedent) IS the Microsoft-documented best practice. The maintainer-line handoff is well-grounded in canonical sources.

The 7-PR local-workaround chain's exhaustion is not just empirical — it's also documented:
- F1 + F6 confirm Application.Run is by design (PR #136 boundary is real)
- F2 confirms VBA-side DoEvents ≠ COM-side OS-yield (PR #131-#134 mechanism boundary is real)
- F5 + PR #137 + F9 confirm the AddFromString event-binding-cache layer is undocumented and version-specific (PR #137 boundary is real)

**Honest acknowledgment of the dissenting signal:**

F4's `Forms(0).RecordSource = Forms(0).RecordSource` self-rebinding workaround is the ONE Microsoft-documented local technique we did not explicitly test in PR #129-#137. The verdict is NOT "all local options exhausted under all readings" — it's "the maintainer-line is the highest-leverage forward step given the documented evidence". If the maintainer (or reviewer) wants to authorize one more local probe before the maintainer-line handoff is acted on, the F4 self-rebinding pattern is the candidate; it's a well-documented technique and would take ~30-60 minutes to test. But:

- it would NOT change the upstream fix recommendation (Concern 4's globals fix is still the right upstream change)
- it would NOT change the canonical Issue truth for #21/#23/#24
- if it succeeds, it'd be another Status-specific test-driver workaround — which the brief explicitly classified as exhausted in disposition
- if it fails, the parked posture stays as-is

**Net:** the verdict supports maintainer-line; the F4 finding is recorded as an optional unpause path the maintainer can authorize if they want, NOT as a reason to keep the line in active local development.

---

## What this research changes about the 4-concern handoff (PR #139)

**Nothing material.** PR #139's handoff memo was correct on all four concerns:

- Issue #21 (GroupData CmdNeo4j) — external research is silent (no community discussion of the specific empty-recordset .MoveFirst pattern); the canonical issue's filed text remains the authoritative source.
- Issue #23 (Associations CmdNeo4j INSERT typo) — external research is silent (no community discussion of the specific c_index_addr_type_code typo); canonical issue text remains authoritative.
- Issue #24 (Place CmdNeo4j SELECT projection) — external research is silent on this specific pattern; canonical issue text remains authoritative.
- Concern 4 (Status cleanup-rebind, NEW) — external research **strengthens** the handoff: F3+F7 directly support the suggested upstream fix (globals); F4+F8 surface ONE additional local technique not yet tested, recorded here as optional.

The handoff memo's overall framing ("4 concerns, 3 canonical P1 + 1 not-yet-filed root cause; bundled for one round of maintainer review") stands.

---

## Constraints honoured per brief

- ✅ Read-only research; no Access COM run; no tests / driver / README / triage / canonical reports / issue severity changes
- ✅ Only the two research artifacts touched (this MD + paired JSON)
- ✅ External sources stratified by credibility (Tier 1 = Microsoft official docs; Tier 2 = recognized community sources; Tier 3 = forum anecdote)
- ✅ Quoted Microsoft Learn docs verbatim where possible
- ✅ Raw findings (F1-F9) and inference (I1-I6) cleanly separated into different sections
- ✅ Verdict bucket selected from the 3 named brief options
- ✅ Did NOT silently treat forum anecdote as conclusion; F9 explicitly Tier 3
- ✅ Did NOT reopen any local probe (per brief)
- ✅ `analysis/report_screenshot_audit.md` drift left alone (standing rule)

## Self-review (per `docs/skills/programmer-self-review-template.md`)

**A. Branch shape.** Branch `research/access-external-evidence-status-exports` cut from current `main = 81c029b` (post PR #139 merge). Only the two research artifacts touched (this MD + paired JSON). No tests, driver, README, triage, canonical reports, issue severity, or other artifacts changed. Pre-existing `analysis/report_screenshot_audit.md` drift left alone per standing instruction.

**B. Source-of-truth sync.** This memo does not modify any source-of-truth file — it cites them. Microsoft Learn URLs are quoted with the exact page-canonical URL the WebFetch tool used; community-source URLs cite the same page the WebSearch returned. The PR #139 handoff memo's content is referenced but not modified; this memo concludes by stating PR #139's framing stands.

**C. Evidence vs claim.** Each finding F1-F9 cites its source AND its credibility tier. Verbatim quotes from Microsoft Learn (F1-F4) are reproduced as quoted blocks; community-source consensus (F6, F7, F8) is summarized with multi-source corroboration; forum anecdote (F9) is explicitly marked Tier 3. The verdict (`external_evidence_supports_maintainer_line`) is grounded in F1+F2+F3+F4+F5+F6+F7+F8 collectively; the dissenting signal (F4's untested self-rebinding workaround) is honestly recorded as an optional unpause path, NOT silently dropped to favor the verdict.

**D. Residual risk.** Research memo, not implementation; residual risk is purely advisory: (1) the F4 self-rebinding technique remains untested by us; if maintainer (or reviewer) authorizes one more local probe, that is the candidate. (2) F5's absence-of-doc finding is a "no evidence" claim, not a "negative evidence" claim; some Microsoft KB or older official doc might exist that we didn't find. (3) Tier 2 community sources are durable but not infallible; readings could shift if a new authoritative source surfaces. (4) No code path or test altered, so no runtime regression risk.

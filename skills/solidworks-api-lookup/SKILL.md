---
name: solidworks-api-lookup
description: Query the local solidworks-api MCP for authoritative SOLIDWORKS API signatures, enum values, return codes, and behavioural contracts before writing, changing, or reviewing VBA in the "VBA 3D TO 2D" project. Use this whenever the work touches a SOLIDWORKS API call, an sw* constant or enum, a drawing/view/dimension/annotation/selection operation, or any file under src/ — including when you are confident you already know the value, because hard-coded constants in this repo have silently been wrong before. Also use when diagnosing why the macro imports the wrong annotations, produces no ordinates, misclassifies views, or returns an unexplained error code.
---

# SOLIDWORKS API lookup for the VBA 3D-to-2D macro

## Why this exists

This macro fails quietly rather than loudly. A wrong bitmask does not raise an error — it imports the wrong annotation category and returns a plausible count. A display-mode constant that names a real but different member does not raise an error — the view just renders in the wrong mode. The cost of a guessed constant is a drawing that looks finished and is wrong.

This actually happened here. The active importer once declared `swInsertDimensionsMarkedForDrawing = 1` and `swInsertHoleWizardCallouts = 64`. Those are really `swInsertCThreads` and `swInsertNotes`. The macro ran, reported success, and imported the wrong things. The correct values (`32768` and `1048576`) came from querying the MCP corpus — see `docs/SOLIDWORKS_API_VALIDATION.md`.

So the rule that matters is simple: **look it up, don't recall it.** `Agents.md` already states the MCP and the installed 2025 type library are authoritative and that enum values must never be guessed or copied from older wrappers. This skill is how to actually do that well.

## Know what the corpus is before you trust it

The MCP serves a **snapshot** of the official SOLIDWORKS API Help, and it does not match your target exactly:

- Its tool metadata cites the SOLIDWORKS **2025** help root, but `docs/SOLIDWORKS_API_VALIDATION.md` records the corpus identifying itself as **2026**, and it is flagged `compatibility-snapshot`.
- Method records carry `since_version: unknown`. The corpus cannot tell you whether a member exists in your build.
- This project targets **SOLIDWORKS 2025 SP1.2**.

That gap is not a reason to skip the MCP — it is by far the fastest way to get an exact signature, enum table, or contract. It is a reason to be precise about what an MCP answer proves:

| The MCP tells you | Confidence |
|---|---|
| A member's name, parameters, and return type | Reliable — API surface rarely churns between adjacent releases |
| Behavioural contract in the **Remarks** section (selection order, required cleanup, coordinate frame) | Reliable — this is the highest-value content in the corpus |
| Exact numeric enum values | Strong evidence, **not proof** for 2025 |
| Whether a member exists in SW2025 | Nothing — verify locally |

Report MCP-derived numbers as **MCP evidence**, and mark anything load-bearing **verify in SW2025** until it is checked against the installed type library / Object Browser. Never present a corpus value as a confirmed 2025 fact.

## Picking the right tool

Six tools. Reaching for the wrong one wastes a round-trip.

| You need | Tool | Notes |
|---|---|---|
| Exact signature, return type, and **contract** for a known member | `solidworks_lookup_method` | Needs `interface` + `member`. Your default. |
| Numeric values for an `sw*_e` enum | `solidworks_get_enum_values` | The single most valuable tool here — bitmasks and error codes. Sanity-check the result (below). |
| To find out *which* interface or member does a thing | `solidworks_search_api` | A locator only. Always confirm the hit with `lookup_method`. |
| Everything available on an interface | `solidworks_get_interface_members` | Good for "is there a proper API for this?" instead of a name-string hack. |
| Adjacent members you might be overlooking | `solidworks_find_related` | Returns bare names only — you must `lookup_method` each one. |
| Which official example topics cover a member | `solidworks_get_examples` | Returns titles and hrefs. **`code_blocks` is usually empty** — expect topic names to open in local SOLIDWORKS help, not runnable code. |

No parameter is marked required in any schema, so a vague call returns something unhelpful rather than erroring. Supply `interface` + `member` whenever you know them.

## Searching effectively

`solidworks_search_api` is keyword-scored, not semantic. This changes how you phrase queries.

Searching `"classify drawing view orientation isometric"` returns mirror-view methods and event-handler delegates — it matched on the bare word "orientation". Searching `"GetVisibleEntities2"` puts the exact method first.

So query in **API vocabulary**, not intent: member-name fragments (`Ordinate`, `InsertModelAnnotations`), interface names (`IView`), or `sw*_e` enum names. If you only know the intent, either search a noun you expect in the member's name, or list the likely interface's members and read down the list. Narrow with `interface` or `categories` (`drawings`, `dimensions`, `annotations`) when a broad query is noisy.

Two things search results give you cheaply, worth scanning for:

- Summaries beginning **"Obsolete. Superseded by …"** — flag these immediately if the macro still calls one.
- **`I`-prefixed twins** (`IGetVisibleEntities2` alongside `GetVisibleEntities2`). Those exist for C++/.NET array marshalling. VBA binds the non-`I` form returning a `System.Object` variant array. Recommending the `I` form to VBA produces code that will not run.

## Reading a lookup result

The signature is the least interesting part. Read in this order:

1. **Remarks** — where the traps live. `AddOrdinateDimension`'s remarks are the reason the datum must be selected first and the reason `IModelDoc2::SetPickMode` is required to close a group. A signature alone would never tell you that.
2. **Return Value** — usually "error as defined by `sw*Error_e`". That is a follow-up `get_enum_values` call, and decoding it is what turns "the call returned 3" into a diagnosis.
3. **Parameters** — note which take an interface (e.g. a `Component2`), because passing `Nothing` where an object is required is an unresolved failure mode in this codebase.
4. **`deprecated`** in the frontmatter, and `related` for alternatives.

Chain the calls: `lookup_method` → `get_enum_values` for every enum the signature names. A method contract is only half-known until its parameter and error enums are decoded.

### When the Remarks and the enum table disagree, the Remarks win

The enum table answers *"what is the value of this member?"* A method's Remarks answer *"which member does **this method** actually want?"* Those are different questions, and the second is nearly always the one you are really asking. An enum is shared across many methods, and any individual method may constrain, ignore, or silently reinterpret part of it.

`IView.SetDisplayMode3` is the standing example. `swDisplayMode_e` lists `swSHADED_EDGES = 7`, so "shaded with edges is 7" reads like a settled fact. But that method's Remarks say shaded-with-edges is obtained by passing `Mode = swSHADED` with edges enabled — and further, that any `swFACETED_*` value passed in `Mode` is silently swapped for its non-faceted equivalent. Take `7` from the enum table alone and you ship a constant that is wrong for the call you are making and raises no error doing it.

So when a method's Remarks name specific members, treat that as the method's real domain and the enum table as background. If you cite a numeric constant for a call, make sure you have read the Remarks of the call — not just the enum it points at.

### Sanity-check the enum table itself

Enum records normally carry the value twice — once in the value field, once in the description as `or 0x20`. That redundancy is worth using, because the parser sometimes mangles it.

`swAutoInsertCenterMarkTypes_e` comes back with all three members reporting `value: "0"` and descriptions reading `"x1 or 1"`, `"x2 or 2"`, `"x4 or 4"`. The leading `0` of `0x1` has been split off into the value column. The real values are 1, 2 and 4.

Two tells: a flag enum in which every member is `0` is impossible, and a description beginning with a bare `x` means the prefix was eaten. When you hit either, the true number is in the description — and say plainly that you reconstructed it, because a reconstructed constant needs the Object Browser check *more* than an ordinary one, not less.

## Turning a lookup into a change

Keep MCP evidence and local verification visibly separate — conflating them is how a 2026 value ends up compiled into a 2025 macro.

**In code**, follow the convention already in `src/active-ordinate/Module4_ModelItemImporter.bas`: name the constant, give its value, and comment the verification source.

```vba
' Verified against the installed SOLIDWORKS 2025 SP1.2 type library.
Private Const swInsertDimensionsMarkedForDrawing As Long = 32768
```

Use that comment only once the value is genuinely confirmed locally. If it is still MCP-only, say so instead — an honest `' MCP corpus value; verify in SW2025 Object Browser before acceptance.` is far better than a comment that overstates what you checked.

**In prose**, cite the member you queried (`IModelDocExtension.AddOrdinateDimension`) and quote the contract you are relying on, so a reviewer can re-run the same lookup.

**In `docs/SOLIDWORKS_API_VALIDATION.md`**, add material findings in the shape that document already uses — an API-contract row stating the MCP contract, the code's actual use, and an assessment; enum values under the relevant subsection; and anything unproven appended to the *Required SW2025 local type-library verification* list. Small confirmations that change nothing do not need recording; a corrected value, a newly discovered contract, or a resolved uncertainty do.

### That document is dated evidence, not current truth

It is the best starting point for what has already been checked, and it is old enough to be wrong about the code. Its finding 5 still states that both ordinate engines omit `SetPickMode`; `src/active-ordinate/Module5_FallbackDimensionEngine.bas` has been calling it for some time. Anyone who takes that line at face value goes looking for a missing call that is already there, and does it with complete confidence — a stale document is more dangerous than no document, because it reads like established fact.

`Agents.md` already settles the precedence: current source wins over an older export, screenshot, transcript, or document. Apply that to this document too. When a finding names a file and a behaviour, confirm the behaviour is still there before you build on it. Where a finding has gone stale, correcting it is usually worth more than the lookup you came for.

## Where this stops

This skill covers establishing what the API actually does. It does not replace `Agents.md`, which governs everything downstream — the protected baseline in `src/baseline-model-dims/`, the three authorized fixtures, ANSI/no-BOM encoding, and the compile–run–inspect loop. An MCP contract is never on its own evidence that a change works: `Agents.md` is explicit that API success is not proof of a correct manufacturing drawing. The lookup tells you the call is *right*; only compiling and running against the fixtures tells you the drawing is.

## Worked example

*"First ordinate group is fine, the second returns 2."*

1. `solidworks_get_enum_values(enum="swCreateOrdDimError_e")` — `2` is `swCreateOrdDimErr_GenNoInternalDims`. Read what that *excludes*: there are dedicated codes for bad selection (3), extra selection (6), duplicate in group (8) and bad direction (9). So SOLIDWORKS accepted the selection and the direction, and generated nothing. That is a "nothing here was dimensionable" result, not a "you called me wrong" result — which already rules out most of the obvious suspects.
2. `solidworks_lookup_method(interface="IModelDocExtension", member="AddOrdinateDimension")` — Remarks give the contract: datum entity selected first, selections after the call keep appending to the open group, `IModelDoc2::SetPickMode` closes it.
3. **Now read the current engine, before concluding anything.** The contract makes a missing `SetPickMode` the tempting story, and `docs/SOLIDWORKS_API_VALIDATION.md` will encourage it — but `src/active-ordinate/Module5_FallbackDimensionEngine.bas` already calls it. Whatever is wrong, it is not that. Skipping this step is how you produce a confident, well-cited, wrong answer.
4. Give ranked candidates with the evidence for each, say plainly which ones you cannot separate from the code alone, and name the one experiment that would distinguish them. An API contract narrows the search; it does not close it.
5. Record what the lookup settled in `docs/SOLIDWORKS_API_VALIDATION.md`, then compile and run against an authorized fixture per `Agents.md` — "the API contract is satisfied" and "the drawing is correct" are different claims.

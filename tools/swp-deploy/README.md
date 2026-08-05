# Target-Spec Hybrid SWP Deployment

This tool promotes the exported source under `src/target-spec-hybrid-v2` into
`Fable.swp` without manually replacing individual modules in the VBA editor.

The deployment is fail-safe:

1. validate the manifest and all source files;
2. inventory the target `.swp` without changing it;
3. require exclusive access to the target;
4. copy the target to timestamped evidence and candidate files;
5. create each managed component with its manifest module/class type and insert
   its metadata-free source into the candidate project;
6. save the already-named candidate through the VBE Save command, then copy it
   to a second candidate file;
7. load the candidate through `ISldWorks.RunMacro2` as a bootstrap execution
   probe;
8. extract all managed VBA source and verify it against the repository; and
9. atomically replace `Fable.swp` only after every prior gate passes.

Manual reference drawings, authorized model files, the protected baseline, and
`src/active-ordinate` are outside the deployment scope.

## One-time bootstrap

`Fable.swp` needs the stable deployment bootstrap once. Future source updates do
not require module-by-module VBA editing.

1. Make or retain a recoverable copy of `Fable.swp`.
2. Open `Fable.swp` with **Tools > Macro > Edit**.
3. In the VBA editor choose **Tools > References** and enable **Microsoft Visual
   Basic for Applications Extensibility 5.3**.
4. Import `tools/swp-deploy/Module0_SourceDeployment.bas` using **File > Import
   File** and confirm its component name is `Module0_SourceDeployment`. If an
   older bootstrap already exists, replace its code with the complete cleaned
   file while retaining that component name.
5. Compile the whole VBA project, save it, and close the VBA editor.

The bootstrap is not part of the managed-component list, so it cannot delete or
replace itself while running. The consequence is that **changes to
`Module0_SourceDeployment.bas` never reach `Fable.swp` through a normal
deploy** — the deploy script only checks that the bootstrap is present. After
editing it, repeat step 4 above.

Step 5's "compile the whole VBA project" is a check, not a precondition. If
the project does not currently compile, import the bootstrap and save anyway,
then deploy — a deploy that prunes orphans is often exactly what is needed to
make the project compile again.

## The UserForms are not deployable either

`UserForm1` and `UserFormSection` are outside the manifest because the
deployer only handles `StdModule` and `ClassModule`. A deploy therefore never
touches them, exactly like the bootstrap: **changing
`src/<trunk>/UserForm1.frm` in the repository does not change the form inside
`Fable.swp`.**

Worse, they cannot be imported. Both `.frm` files start at
`Attribute VB_Name` with no `VERSION 5.00` / `Begin ... End` designer block,
because both forms create their controls at runtime rather than from a
designer layout. There is no `.frx`. **File > Import File will not
reconstruct these forms.**

To change a form, open the existing component in the VBE and replace its code
in place, keeping the component name — the same procedure as the bootstrap.
Strip the leading `Attribute` lines from the repository copy first; the VBE
supplies its own.

This bit on 2026-08-05: after the trunk change and the orphan prune, the
project still failed to compile because the embedded `UserForm1` was the old
implementation calling `Module1_Main.GetFixtureKey`, a member the new trunk
does not define.

## Pruning components that left the manifest

The deployer replaces the components the manifest names. Until 2026-08-05 it
did nothing about components that were embedded but no longer named, so
shrinking the manifest stranded them: the 38-to-12 change orphaned 26 modules
that still called members the new source did not define, and the project
stopped compiling.

`Deploy-TargetSpecHybrid.ps1` now writes `PRUNE_UNMANAGED=1` plus a
`PROTECTED=` line per component that must survive, and the bootstrap removes
every standard or class module that is neither managed nor protected. Each
removal is logged as `PRUNE|name=...|status=REMOVING`, followed by
`PRUNE|status=COMPLETE|removed=<n>`, in the deployment result file.

Three guards, because this deletes code:

1. Only standard and class modules are eligible. Forms and document classes
   are a different component type and are never removed.
2. The protected list carries the bootstrap's own name, so the prune cannot
   remove the module executing it.
3. `ThisLibrary`, `UserForm1` and `UserFormSection` are protected explicitly,
   because they sit outside the manifest by design and are otherwise
   indistinguishable from an orphan by name.

Deployable `.bas` and ordinary `.cls` files are Windows-1252/ANSI without a BOM
and contain no VBA `Attribute` metadata. The bootstrap creates each component
with the manifest's explicit standard-module or class-module type, assigns its
name, and inserts its code. UserForm and `ThisLibrary` designer metadata are
outside the managed deployment scope.

## Preflight

From PowerShell in the workspace:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\swp-deploy\Deploy-TargetSpecHybrid.ps1 -PreflightOnly
```

Preflight is read-only. It reports the target, expected source revision,
managed-component count, and whether the one-time bootstrap is installed.

## Deploy

Close `Fable.swp` in the VBA editor but leave SOLIDWORKS running. Then execute:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\swp-deploy\Deploy-TargetSpecHybrid.ps1
```

Evidence and both pre-deployment copies are retained under:

```text
test_assets/iteration_evidence/swp_deployment/<timestamp>/
```

If any import, save, compile, read-back, revision, or file-lock check fails, the
original `Fable.swp` is not replaced.

`COMPILE_PROBE|status=SUCCESS` is the historical output name of the bootstrap
execution probe. It proves that `Module0_SourceDeployment.CompileProbe` could
run, but it does **not** execute VBA Editor **Compile Project** and can miss a
compile error in an unexecuted managed module. A manual full-project compile
remains mandatory after every deployment.

## Managed scope

The manifest deploys all nine standard production modules and the six ordinary
class modules. It deliberately leaves these host/designer components alone:

- `ThisLibrary`;
- `UserForm1`;
- `UserFormSection`.

The current forms are code snapshots rather than native `.frm/.frx` designer
exports, so automatically rebuilding their designers would not be safe. If a
future demonstrated defect requires a form or `ThisLibrary` change, treat that
as a separate deployment extension.

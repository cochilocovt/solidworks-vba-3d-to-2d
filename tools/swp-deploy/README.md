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
7. load the candidate through `ISldWorks.RunMacro2` as a compile probe;
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
replace itself while running.

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

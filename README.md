# Bulk RVT to NWC Exporter

A Revit add-in that scans a selected directory and exports every discovered Revit project file to Navisworks cache format (.nwc).

## Compatibility

The checked-in project targets .NET Framework 4.8 and references the Revit 2022 API assemblies. It is a Revit add-in, not a standalone batch utility. Other Revit versions require compatible API references and validation.

## Features

- Select a directory containing Revit project files.
- Recursively scan the selected folder and its subfolders for .rvt files.
- Open each discovered file through the Revit application API.
- Export the whole model using NavisworksExportScope.Model.
- Save the NWC output beside the source RVT with the same base filename.
- Close each opened document without saving changes.
- Display a completion message, or an error message if an exception interrupts the run.

## Workflow and options

The current UI exposes a folder selection dialog. It does not expose filters, custom export scope, output directory selection, naming rules, or a progress/cancel control.

1. Open Revit 2022 and a host project.
2. Launch the NWC export command.
3. Select the directory containing RVT files.
4. The add-in searches that directory and all subdirectories.
5. Each discovered RVT is opened, exported to NWC in the same directory, and closed without saving.
6. Review the completion message and verify the generated NWC files.

**Run on copies first.** Existing NWC files with the same name may be replaced by the export workflow; verify output paths and preserve deliverables before running.

## Build from source

1. Install Visual Studio with .NET desktop development and the .NET Framework 4.8 targeting pack.
2. Install Revit 2022 and verify that the Navisworks export functionality is available.
3. Open NWCExporter/NWCExporter.sln.
4. Check the RevitAPI.dll and RevitAPIUI.dll paths in NWCExporter.csproj and update them to the installed Revit 2022 directory.
5. Build the solution.
6. Locate NWCExporter.dll and configure a Revit .addin manifest whose assembly path points to the built DLL and whose command class matches the source.

## Install

1. Close Revit.
2. Place NWCExporter.dll in a stable folder, such as a dedicated subfolder under %APPDATA%\Autodesk\Revit\Addins\2022.
3. Place a correctly configured .addin manifest in %APPDATA%\Autodesk\Revit\Addins\2022.
4. Verify the assembly path and external command class in the manifest.
5. Restart Revit and confirm the command is available.

## Troubleshooting

- **Command not visible:** check the .addin manifest, assembly path, command class, and version-specific folder.
- **No files found:** confirm the selected directory contains accessible .rvt files; the scan includes subdirectories.
- **Export fails:** manually test a Navisworks export from the same Revit installation and verify the exporter integration.
- **File access errors:** check file locks, network access, and permissions.
- **Unexpected results or interruption:** the current implementation catches exceptions at the overall export workflow level; test smaller folders to isolate problematic models.

## Limitations

- No per-file progress window, result table, or configurable output location is exposed by the current implementation.
- Scans subfolders automatically; there is no UI option to disable recursion.
- Exports the model scope and uses the source RVT directory for the NWC output.
- Validate each output before relying on it for coordination or contractual deliverables.

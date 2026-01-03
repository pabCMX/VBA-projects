Attribute VB_Name = "DEPRECATED_FILES"
' ============================================================================
' DEPRECATED_FILES - Documentation of Files NOT Migrated to V2
' ============================================================================
' WHAT THIS MODULE IS:
'   This is a documentation-only module that lists files from the original
'   populateMacro/ directory that were intentionally NOT migrated to V2.
'   
'   Each file is listed with the reason it was deprecated and what (if
'   anything) replaced its functionality.
'
' WHY SOME FILES WEREN'T MIGRATED:
'   1. Development/Debug utilities - Not production code
'   2. Obsolete technology - Internet Explorer automation
'   3. Replaced by better solutions - Native VBA instead of custom classes
'   4. Empty/Dead code - Never executed or 100% commented out
'   5. Trivial wrappers - Inlined into calling code
'
' IMPORTANT:
'   If you're looking for functionality from one of these files, check the
'   "REPLACED BY" section to find where that functionality now lives.
'
' ============================================================================

Option Explicit

' ============================================================================
' DEPRECATED: IE_Close.vba
' ============================================================================
' WHAT IT DID:
'   Windows API calls to find and close Internet Explorer windows.
'   Used PostMessage to send WM_CLOSE to browser windows.
'
' WHY DEPRECATED:
'   Internet Explorer was deprecated by Microsoft and removed from Windows.
'   This code is no longer needed as BIS doesn't use IE anymore.
'
' REPLACED BY:
'   Nothing - functionality is obsolete.
' ============================================================================


' ============================================================================
' DEPRECATED: ExportAllMacros.vba / Module5.vba (duplicate)
' ============================================================================
' WHAT IT DID:
'   Development utility that exported all VBA modules from the workbook
'   to .vba files on disk. Used for version control and backup.
'
' WHY DEPRECATED:
'   This was a development tool, not production code. It was used by
'   developers to extract code for editing/backup, not by examiners.
'
' REPLACED BY:
'   Nothing in V2 - development utilities are kept separate from
'   production code. Use VBA IDE export or a Git tool for VBA.
' ============================================================================


' ============================================================================
' DEPRECATED: Module2.vba
' ============================================================================
' WHAT IT DID:
'   Contained CopyModule() and CopyForm() functions for copying VBA
'   components between workbooks using temp file export/import.
'
' WHY DEPRECATED:
'   Functionality was moved inline to Pop_Repopulate.vba with V2
'   comments explaining the VBProject import/export process.
'
' REPLACED BY:
'   Pop_Repopulate.vba - same functionality with better documentation.
' ============================================================================


' ============================================================================
' DEPRECATED: Module4.vba
' ============================================================================
' WHAT IT DID:
'   - Find_Names(): Debugging utility to list named ranges on active sheet
'   - group_box_outline_remove(): Hid group box outlines on a sheet
'
' WHY DEPRECATED:
'   Both were debugging/development utilities, not production features.
'
' REPLACED BY:
'   Nothing - not needed in production.
' ============================================================================


' ============================================================================
' DEPRECATED: Module5.vba
' ============================================================================
' WHAT IT DID:
'   Duplicate of ExportAllMacros.vba (same code in two files).
'
' WHY DEPRECATED:
'   Duplicate code, and development utility (see ExportAllMacros above).
'
' REPLACED BY:
'   Nothing - see ExportAllMacros.
' ============================================================================


' ============================================================================
' DEPRECATED: redisplayform1_mod.vba
' ============================================================================
' WHAT IT DID:
'   Contained a single 3-line subroutine that showed UserForm50.
'   Just called: UserForm50.Show
'
' WHY DEPRECATED:
'   Trivial wrapper with no logic. Inlined into the calling code.
'
' REPLACED BY:
'   Direct calls to UF_PopulateMain.Show where needed.
' ============================================================================


' ============================================================================
' DEPRECATED: clFileSearchModule.vba
' ============================================================================
' WHAT IT DID:
'   Custom file search class (~1000 lines) using Windows API to search
'   for files recursively. Supported wildcards and depth limits.
'
' WHY DEPRECATED:
'   Over-engineered for the use case. VBA's native Dir() function
'   with a simple recursive wrapper does the same thing in 50 lines.
'
' REPLACED BY:
'   Pop_Repopulate.vba uses native Dir() with FindFileInFolder() helper.
'   Common_Utils.PathExists() for simple path validation.
' ============================================================================


' ============================================================================
' DEPRECATED: UserForm3.frm
' ============================================================================
' WHAT IT DID:
'   Displayed a form asking the user to manually select which drive
'   letter contained their DQC folder (E:, F:, G:, etc.)
'
' WHY DEPRECATED:
'   No longer needed - Common_Utils.GetDQCDriveLetter() now auto-detects
'   the network drive by checking for known UNC paths.
'
' REPLACED BY:
'   Common_Utils.GetDQCDriveLetter() - automatic drive detection.
' ============================================================================


' ============================================================================
' DEPRECATED: Sheet1.vba
' ============================================================================
' WHAT IT DID:
'   Empty sheet module - no code at all (just VB attributes).
'
' WHY DEPRECATED:
'   No functionality to migrate.
'
' REPLACED BY:
'   Nothing - was always empty.
' ============================================================================


' ============================================================================
' DEPRECATED: Sheet6.vba
' ============================================================================
' WHAT IT DID:
'   100% commented out code. Originally synced data between PA721 sheet
'   and other parts of the workbook.
'
' WHY DEPRECATED:
'   All code was already commented out in the original file.
'   Appears to have been disabled years ago but never removed.
'
' REPLACED BY:
'   Nothing - functionality was already disabled.
' ============================================================================


' ============================================================================
' DEPRECATED: Sheet9.vba
' ============================================================================
' WHAT IT DID:
'   - shapes_outline(): Debug utility to list all shapes on a sheet
'   - clear_buttons(): Cleared button values for testing
'
' WHY DEPRECATED:
'   Debugging utilities, not production code.
'
' REPLACED BY:
'   Nothing - development tools.
' ============================================================================


' ============================================================================
' DEPRECATED: Sheet28.vba
' ============================================================================
' WHAT IT DID:
'   Empty sheet module - no code at all (just VB attributes).
'
' WHY DEPRECATED:
'   No functionality to migrate.
'
' REPLACED BY:
'   Nothing - was always empty.
' ============================================================================


' ============================================================================
' MODULE MAPPING REFERENCE
' ============================================================================
' Quick reference showing what V2 modules replaced V1 modules:
'
' V1 File                         -> V2 File
' ============================================================================
' populate_mod.vba               -> Pop_Main.vba
' Populate_snap_pos_delimited_mod.vba -> Pop_SNAP_Positive.vba
' populate_snap_neg_delimited_mod.vba -> Pop_SNAP_Negative.vba
' Populate_TANF_delimited_mod.vba     -> Pop_TANF.vba
' Populate_MA_delimited_mod.vba       -> Pop_MA.vba
' repopulate_mod.vba             -> Pop_Repopulate.vba
' TransPopulate.vba              -> Pop_Transmittals.vba
' Module1.vba                    -> Review_Approval.vba
' Module3.vba                    -> Review_EditCheck.vba
' Drop.vba                       -> Review_Drop.vba
' TANFmod.vba                    -> Review_TANF_Utils.vba
' Module11.vba                   -> Review_SNAP_Utils.vba
' MA_Comp_mod.vba                -> Review_MA_Comp.vba
' GAGetElements.vba              -> Review_GA_Elements.vba
' Finding_Memo.vba               -> Review_FindingMemo.vba
' CashMemos.vba                  -> Review_CashMemos.vba
' CAO_Appointment.vba            -> Review_Appointments.vba
' ThisWorkbook.vba               -> Review_ValidationHooks.vba
' Sheet4/8/12/14/18/20/22/25/27  -> Review_SheetEvents.vba (consolidated)
' UserForm50.frm                 -> UF_PopulateMain.frm
' UserForm1.frm                  -> UF_TANF_ResultsColumn.frm
' UserForm2.frm                  -> UF_TANF_FinalDetermination.frm
' SelectForms.frm                -> UF_SelectForms.frm
' SelectDate.frm                 -> UF_DatePicker.frm
' SelectTime.frm                 -> UF_TimePicker.frm
' MASelectForms.frm              -> UF_MA_SelectForms.frm
' UserFormMAC2.frm               -> UF_MA_Comp2.frm
' UserFormMAC3.frm               -> UF_MA_Comp3.frm
' GAUserForm1.frm                -> UF_GA_Helper1.frm
' GAUserForm2.frm                -> UF_GA_Helper2.frm
' (new)                          -> Common_Utils.vba (consolidated utilities)
' (new)                          -> Config_Settings.vba (centralized constants)
' ============================================================================



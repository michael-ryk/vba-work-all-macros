Sub UpdateMainModule()
  '
  ' Update Macros
  '
  'Definition
  Set VBProj = Workbooks("Personal.xlsb").VBProject

  On Error GoTo ErrModuleNotExist
  Set VBComp = VBProj.VBComponents("Module_MichaelR_Macros")

  'Delete existing component
  VBProj.VBComponents.Remove VBComp
  Debug.Print ("Module Removed")

  'Add updated component
  ' VBProj.VBComponents.Import "c:\tmp\TestModule.bas" 'For Debug use only
  On Error GoTo ErrOnImport
  VBProj.VBComponents.Import "\\emcsrv\R&D\r&d_work_space\Teams\Validation & Verification\Hadar&Meira\Alex_H_team\VBA-Script-Report-Analyzing\Module_MichaelR_Macros.bas"
  MsgBox ("Import Success!"), , "Import Successfully"

  Exit Sub

      ' Error Handling
  ErrOnImport:
      MsgBox ("Sorry, but now you without macro at all. Module removed but failed to import. Please check Connection to emcsrv available. "), vbCritical, "IMPORT Failed"
  Exit Sub
      
  ErrModuleNotExist:
      MsgBox ("Unable to Remove module - it not exist."), vbExclamation, "Module not exist"
  Resume Next

End Sub
Attribute VB_Name = "errdemo"
Option Explicit

Sub errdemo()
  lp False
  lp True
  lp False
  lp True
End Sub

Private Sub lp(force_error As Boolean)
  Dim lp As LongPtr
  Dim smcp As glp_smcp
  Dim ret As Long
  Dim name() As Byte
  Dim ind(2) As Long
  Dim val(2) As Double
  
  On Error GoTo error0
  
  ' Register error hook function
  glp_error_hook AddressOf error_hook
  
  ' Register terminal hook function
  glp_term_hook AddressOf term_hook
  
  ' Create problem
  lp = glp_create_prob()
  name = str2bytes("Linear Problem")
  glp_set_prob_name lp, name(0)
  
  If force_error Then
    glp_add_cols lp, -1
  End If
  
  ' Create columns
  glp_add_cols lp, 1
  name = str2bytes("x1")
  glp_set_col_name lp, 1, name(0)
  glp_set_col_kind lp, 1, GLP_CV
  glp_set_col_bnds lp, 1, GLP_UP, 0#, 1#
  
  ' Define objective
  name = str2bytes("obj")
  glp_set_obj_name lp, name(0)
  glp_set_obj_dir lp, GLP_MAX
  glp_set_obj_coef lp, 1, 1

  ' Solve model
  ret = glp_init_smcp(smcp)
  ret = glp_simplex(lp, smcp)
  
  'Retrieve solution
  If ret = 0 Then
    write_lp_solution (lp)
  End If
  glp_delete_prob lp
  
  ' Deregister terminal hook function
  glp_term_hook 0, 0
  
  Exit Sub

error0:
  Debug.Print "Error in GLPK library has been caught"
  Debug.Print Err.Description
  
  On Error GoTo 0
  ' Register error hook function
  glp_error_hook 0, 0
  
End Sub







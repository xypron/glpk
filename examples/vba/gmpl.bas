Attribute VB_Name = "gmpl"

Public Sub gmpl_main(fname As String)
  Dim lp As LongPtr
  Dim glp_tran As LongPtr
  Dim iocp As glp_iocp
  Dim name() As Byte
  
  Dim skip As Long
  Dim ret As Long
  
  ' Register error hook function
  glp_error_hook AddressOf error_hook
  
  ' Register terminal hook function
  glp_term_hook AddressOf term_hook
  
  lp = glp_create_prob()
  
  tran = glp_mpl_alloc_wksp()
  
  name = str2bytes(fname)
  ret = glp_mpl_read_model(tran, name(0), skip)
  If ret <> 0 Then
      glp_mpl_free_wksp (tran)
    glp_delete_prob (lp)
    Err.Raise 1004, "Model file not valid: " & fname
  End If
  
  ret = glp_mpl_generate(tran, 0)
  If ret <> 0 Then
    glp_mpl_free_wksp (tran)
    glp_delete_prob (lp)
    Err.Raise 1004, "Cannot generate model: " & fname
  End If
  
  ' build model
  glp_mpl_build_prob tran, lp

  ' set solver parameters
  glp_init_iocp iocp
  iocp.presolve = GLP_ON
  
  ' solve model
  ret = glp_intopt(lp, iocp)

  ' postsolve model
  If ret = 0 Then
    ret = glp_mpl_postsolve(tran, lp, GLP_MIP)
  End If
  
  ' Free memory
    glp_mpl_free_wksp (tran)
    glp_delete_prob lp
  
  ' Deregister error hook function
  glp_error_hook 0, 0

  Exit Sub

error0:
  Debug.Print Format(Err.Number, "0 - "); Err.Description

  If Err.Number <> GLPK_LIB_ERROR Then
    On Error GoTo 0
    Resume
  End If
End Sub

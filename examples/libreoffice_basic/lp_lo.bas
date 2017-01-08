REM  *****  BASIC  *****

' lp.bas
'
' Copyright (C) 2017, Heinrich Schuchardt <xypron.glpk@gmx.de>.
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
' You should have received a copy of the GNU General Public License
' along with GLPK. If not, see <http://www.gnu.org/licenses/>.

Option Explicit

Public Const GLP_CV = 1              ' continuous variable
Public Const GLP_DB = 4              ' double-bounded variable
Public Const GLP_MIN = 1             ' minimization
Public Const GLP_UP = 3              ' variable with upper bound

Public Type glp_smcp
  ' simplex method control parameters
  msg_lev As Long                    ' message level:
  meth As Long                       ' simplex method option:
  pricing As Long                    ' pricing technique:
  r_test As Long                     ' ratio test technique:
  tol_bnd As Double                  ' spx.tol_bnd
  tol_dj As Double                   ' spx.tol_dj
  tol_piv As Double                  ' spx.tol_piv
  obj_ll As Double                   ' spx.obj_ll
  obj_ul As Double                   ' spx.obj_ul
  it_lim As Long                     ' spx.it_lim
  tm_lim As Long                     ' spx.tm_lim (milliseconds)
  out_frq As Long                    ' spx.out_frq
  out_dly As Long                    ' spx.out_dly (milliseconds)
  presolve As Long                   ' enable/disable using LP presolver
  align_1 As Long                    ' only used for alignment
  foo_bar(35) As Double              ' (reserved)
End Type

Public Type tind
  ind(2) As Long
End Type

Public Type tval
  val(2) As Double
End Type

Declare Function glp_create_prob Lib "glpk" () As Long
Declare Sub glp_set_prob_name Lib "glpk" (ByVal lp As Long, ByVal name As String)
Declare Function glp_add_cols Lib "glpk" (ByVal lp As Long, ByVal count As Long) As Long
Declare Sub glp_set_col_name Lib "glpk" (ByVal lp As Long, ByVal col As Long, ByVal name As String)
Declare Sub glp_set_col_bnds Lib "glpk" (ByVal lp As Long, ByVal row As Long, ByVal typ As Long, ByVal lb As Double, ByVal ub As Double)
Declare Sub glp_set_col_kind Lib "glpk" (ByVal lp As Long, ByVal j As Long, ByVal kind As Long)
Declare Function glp_add_rows Lib "glpk" (ByVal lp As Long, ByVal count As Long) As Long
Declare Sub glp_set_row_name Lib "glpk" (ByVal lp As Long, ByVal row As Long, ByVal name As String)
Declare Sub glp_set_mat_row Lib "glpk" (ByVal lp As Long, ByVal i As Long, ByVal length As Long, ByRef ind As tind, ByRef val As tval)
Declare Sub glp_set_row_bnds Lib "glpk" (ByVal lp As Long, ByVal row As Long, ByVal typ As Long, ByVal lb As Double, ByVal ub As Double)
Declare Sub glp_set_obj_name Lib "glpk" (ByVal lp As Long, ByVal name As String)
Declare Sub glp_set_obj_dir Lib "glpk" (ByVal lp As Long, ByVal dir As Long)
Declare Sub glp_set_obj_coef Lib "glpk" (ByVal lp As Long, ByVal col As Long, ByVal val As Double)
Declare Function glp_init_smcp Lib "glpk" (ByRef smcp As glp_smcp) As Long
Declare Function glp_simplex Lib "glpk" (ByVal lp As Long, ByRef smcp As glp_smcp) As Long
Declare Sub glp_delete_prob Lib "glpk" (ByVal lp As Long)
Declare Function glp_version Lib "glpk" () As String
Declare Function glp_get_prob_name Lib "glpk" (ByVal lp As Long) As String
Declare Function glp_get_obj_name Lib "glpk" (ByVal lp As Long) As String
Declare Function glp_get_obj_val Lib "glpk" (ByVal lp As Long) As Double
Declare Function glp_get_num_cols Lib "glpk" (ByVal lp As Long) As Long
Declare Function glp_get_col_name Lib "glpk" (ByVal lp As Long, ByVal j As Long) As String
Declare Function glp_get_col_prim Lib "glpk" (ByVal lp As Long, ByVal j As Long) As Double


'  Minimize z = -.5 * x1 + .5 * x2 - x3 + 1
'
'  subject to
'  0.0 <= x1 - .5 * x2 <= 0.2
'  -x2 + x3 <= 0.4
'  where,
'  0.0 <= x1 <= 0.5
'  0.0 <= x2 <= 0.5
'  0.0 <= x3 <= 0.5
Sub lp()
  Dim lp As Long
  Dim smcp As glp_smcp
  Dim ret As Long
  Dim name() As Byte
  Dim ind As tind
  Dim val As tval
  
  ' Create problem
  lp = glp_create_prob()
  glp_set_prob_name lp, "Linear Problem"
  
  ' Create columns
  glp_add_cols lp, 3
  glp_set_col_name lp, 1, "x1"
  glp_set_col_kind lp, 1, GLP_CV
  glp_set_col_bnds lp, 1, GLP_DB, 0#, 0.5
  glp_set_col_name lp, 2, "x2"
  glp_set_col_kind lp, 2, GLP_CV
  glp_set_col_bnds lp, 2, GLP_DB, 0#, 0.5
  glp_set_col_name lp, 3, "x3"
  glp_set_col_kind lp, 3, GLP_CV
  glp_set_col_bnds lp, 3, GLP_DB, 0#, 0.5
  
  ' Create rows
  glp_add_rows lp, 2
  
  glp_set_row_name lp, 1, "c1"
  glp_set_row_bnds lp, 1, GLP_DB, 0, 0.2

  ind.ind(1) = 1
  ind.ind(2) = 2
  val.val(1) = 1#
  val.val(2) = -0.5
  glp_set_mat_row lp, 1, 2, ind, val

  glp_set_row_name lp, 2, "c2"
  glp_set_row_bnds lp, 2, GLP_UP, 0, 0.4
  
  ind.ind(1) = 2
  ind.ind(2) = 3
  val.val(1) = -1
  val.val(2) = 1
  glp_set_mat_row lp, 2, 2, ind(0), val(0)
  
  ' Define objective
  glp_set_obj_name lp, "obj"
  glp_set_obj_dir lp, GLP_MIN
  glp_set_obj_coef lp, 0, 1#
  glp_set_obj_coef lp, 1, -0.5
  glp_set_obj_coef lp, 2, 0.5
  glp_set_obj_coef lp, 3, -1

  ' Write model to file
  ' name = str2bytes("c:\temp\lp.lp")
  ' ret = glp_write_lp(lp, 0, name(0))

  ' Solve model
  ret = glp_init_smcp(smcp)
  ret = glp_simplex(lp, smcp)
  
  'Retrieve solution
  If ret = 0 Then
    write_lp_solution (lp)
  End If

  ' Free memory
  glp_delete_prob lp
End Sub

Private Sub write_lp_solution(lp As Long)
  Dim i, n As Long
  Dim name As String
  Dim val As Double
  dim txt as String
        
  name = glp_version()
  txt = txt & "GLPK " & name & Chr$(13) & Chr$(10)
  
  name = glp_get_prob_name(lp)
  txt = txt & "Solution of " & name & Chr$(13) & Chr$(10)
          
  name = glp_get_obj_name(lp)
  val = glp_get_obj_val(lp)
  txt = txt & name & " = " & val & Chr$(13) & Chr$(10)
          
  n = glp_get_num_cols(lp)
  For i = 1 To n
     name = glp_get_col_name(lp, i)
     val = glp_get_col_prim(lp, i)
     txt = txt & name & " = " & val & Chr$(13) & Chr$(10)
   Next i
  MsgBox txt 
End Sub



Attribute VB_Name = "EAN13"

Function FormatConvertEAN13(Any_r As Range) As String
Dim sval As String
Dim frml As String

With Any_r
    sval = CStr(.Value)
    
    frml = "=CONCATENATE(VALUE(MID(" & sval & ",1,1)),CHAR(VALUE(MID(" & sval & ",2,1))+65),IF(VALUE(MID(" & sval & ",1,1))<4,CHAR(VALUE(MID(" & sval & ",3,1))+65),CHAR(VALUE(MID(" & sval & ",3,1))+75)),IF(OR(VALUE(MID(" & sval & ",1,1))=0,VALUE(MID(" & sval & ",1,1))=4,VALUE(MID(" & sval & ",1,1))=7,VALUE(MID(" & sval & ",1,1))=8),CHAR(VALUE(MID(" & sval & ",4,1))+65),CHAR(VALUE(MID(" & sval & ",4,1))+75)),IF(OR(VALUE(MID(" & sval & ",1,1))=0,VALU" & _
            "E(MID(" & sval & ",1,1))=1,VALUE(MID(" & sval & ",1,1))=4,VALUE(MID(" & sval & ",1,1))=5,VALUE(MID(" & sval & ",1,1))=9),CHAR(VALUE(MID(" & sval & ",5,1))+65),CHAR(VALUE(MID(" & sval & ",5,1))+75)),IF(OR(VALUE(MID(" & sval & ",1,1))=0,VALUE(MID(" & sval & ",1,1))=2,VALUE(MID(" & sval & ",1,1))=5,VALUE(MID(" & sval & ",1,1))=6,VALUE(MID(" & sval & ",1,1))=7),CHAR(VALUE(MID(" & sval & ",6,1))+65),CHAR(VALUE(MID(" & sval & ",6,1))+75)),IF(OR(VAL" & _
            "UE(MID(" & sval & ",1,1))=0,VALUE(MID(" & sval & ",1,1))=3,VALUE(MID(" & sval & ",1,1))=6,VALUE(MID(" & sval & ",1,1))=8,VALUE(MID(" & sval & ",1,1))=9),CHAR(VALUE(MID(" & sval & ",7,1))+65),CHAR(VALUE(MID(" & sval & ",7,1))+75)),""*"",CHAR(VALUE(MID(" & sval & ",8,1))+97),CHAR(VALUE(MID(" & sval & ",9,1))+97),CHAR(VALUE(MID(" & sval & ",10,1))+97),CHAR(VALUE(MID(" & sval & ",11,1))+97),CHAR(VALUE(MID(" & sval & ",12,1))+97),CHAR(VALUE(" & _
            "MID(" & sval & ",13,1))+97),""+"")"
    
    On Error GoTo Except
    .FormulaR1C1 = frml
    .Font.Name = "Code EAN13"
    .Font.Size = 34
    .Value = .Value
End With

Except:
    If Err.Number Then
        Any_r.Value = Any_r.Value
    End If
End Function


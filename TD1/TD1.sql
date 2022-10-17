Private Sub Cmd_lecture_Click()
    '1- déclaration des variables'
    Dim db As DAO.Database
    Dim tb_agence As DAO.Recordset
    '2- Affectation'
    Set db = CurrentDb()
    Set tb_agence = db.OpenRecordset("SELECT * FROM AGENCE", dbOpenDynaset)
    If Not tb_agence.EOF Then
        tb_agence.MoveFirst
        tb_agence.MoveNext
        MsgBox tb_agence("VILLE")
    End If
End Sub

Private Sub Cmd_insertion_DblClick(Cancel As Integer)
    '1- déclaration des variables'
    Dim db As DAO.Database
    Dim tb_agence As DAO.Recordset
    '2- Affectation'
    Set db = CurrentDb()
    Set tb_agence = db.OpenRecordset("AGENCE", dbOpenDynaset)
    
    tb_agence.AddNew
    
    tb_agence("CODE") = InputBox("Rentre un code")
    tb_agence("VILLE") = InputBox("Rentre un nom de ville")
    
    tb_agence.Update
End Sub

Private Sub Cmd_modification_GotFocus()
    '1- déclaration des variables'
    Dim db As DAO.Database
    Dim tb_agence As DAO.Recordset
    '2- Affectation'
    Set db = CurrentDb()
    Set tb_agence = db.OpenRecordset("Select * from AGENCE where [VILLE] = 'Tours'", dbOpenDynaset)
    If Not tb_agence.EOF Then
        MsgBox "trouvé"
        tb_agence.Edit
        tb_agence("VILLE") = "Poitiers"
        tb_agence.Update
    End If
End Sub

Private Sub Cmd_supression_Click()
    '1- déclaration des variables'
    Dim db As DAO.Database
    Dim tb_agence As DAO.Recordset
    '2- Affectation'
    Set db = CurrentDb()
    Set tb_agence = db.OpenRecordset("SELECT * FROM AGENCE WHERE [VILLE] = 'Poitiers'", dbOpenDynaset)
    If Not tb_agence.EOF Then
        tb_agence.Delete
        tb_agence.Update
        tb_agence.MoveNext
    End If
End Sub

--PARTIE 2
Option Compare Database
Dim cn As ADODB.Connection
Dim rs AS New ADODB.RecordSet

Private Sub From_Open(Cancel As Integer)
    Set objConn = CurrentProject.Connection
    'rs.Open "select * from AVION", objConn, adOpenKeySet, adLockOptimistic' "Meme chose que la ligne en dessous"
    rs.Open "R1_TD2", objConn, adOpenKeySet, adLockOptimistic
End Sub

Private Sub Debut_Click()
    rs.MoveFirst
    If Not rs.EOF Then
        etiquette.Value = "L'avion " & rs.Fields("AVNOM").Value & " a une capacité de : " & rs.Fields("CAPACITE").Value
    End If
End Sub

Private Sub Precedent_Click()
    If Not rs.BOF Then
        rs.MovePrevious
        If Not rs.BOF Then
            etiquette.Value = "L'avion " & rs.Fields("AVNOM").Value & " a une capacité de : " & rs.Fields("CAPACITE").Value
        Else
            etiquette.Value = "Pas d'avion"
        End If
    End If
End Sub

Private Sub Suivant_click()
    If Not rs.EOF Then
        rs.MoveNext
        If Not rs.EOF Then
            etiquette.Value = "L'avion " & rs.Fields("AVNOM").Value & " a une capacité de : " & rs.Fields("CAPACITE").Value
        Else
            etiquette.Value = "Pas d'avion"
        End If
    End If
End Sub

Private Sub Fin_Click()
    If Not rs.EOF Then
        rs.MoveLast
        If Not rs.EOF Then
            etiquette.Value = "L'avion " & rs.Fields("AVNOM").Value & " a une capacité de : " & rs.Fields("CAPACITE").Value
        Else
            etiquette.Value = "Pas d'avion"
        End If
    End If
End Sub

--Partie3


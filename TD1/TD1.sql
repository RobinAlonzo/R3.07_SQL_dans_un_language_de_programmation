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

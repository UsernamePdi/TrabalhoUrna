On Error Resume Next

Dim fso, scriptDir, dbPath, accessApp, db

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
dbPath = scriptDir & "\UrnaEletronica.accdb"

If fso.FileExists(dbPath) Then
    fso.DeleteFile(dbPath)
End If

Set accessApp = CreateObject("Access.Application")
accessApp.Visible = False
accessApp.NewCurrentDatabase dbPath

Set db = accessApp.CurrentDb

' Criação das tabelas
db.Execute "CREATE TABLE tbl_Cargos (ID_Cargo COUNTER PRIMARY KEY, Nome_Cargo VARCHAR(50), Qtd_Digitos INT);"
db.Execute "CREATE TABLE tbl_Partidos (ID_Partido COUNTER PRIMARY KEY, Numero_Partido VARCHAR(5), Nome_Partido VARCHAR(100), Sigla VARCHAR(10));"
db.Execute "CREATE TABLE tbl_Eleitores (ID_Eleitor COUNTER PRIMARY KEY, Titulo_Eleitoral VARCHAR(12), Nome_Eleitor VARCHAR(100), CPF VARCHAR(14), Zona VARCHAR(4), Secao VARCHAR(4), Ja_Votou YESNO);"
db.Execute "CREATE TABLE tbl_Candidatos (ID_Candidato COUNTER PRIMARY KEY, Numero_Candidato VARCHAR(10), Nome_Candidato VARCHAR(100), Nome_Vice VARCHAR(100), ID_Cargo INT, ID_Partido INT, Caminho_Foto VARCHAR(255));"
db.Execute "CREATE TABLE tbl_Votos (ID_Voto COUNTER PRIMARY KEY, ID_Cargo INT, ID_Candidato INT, Tipo_Voto VARCHAR(20), DataHora_Voto DATETIME);"

accessApp.Quit

If fso.FileExists(dbPath) Then
    MsgBox "O banco de dados 'UrnaEletronica.accdb' foi criado com sucesso!", 64, "Sucesso"
Else
    MsgBox "Erro ao criar o banco. Certifique-se de que o Microsoft Access está instalado.", 48, "Atenção"
End If
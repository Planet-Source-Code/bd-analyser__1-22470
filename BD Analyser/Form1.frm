VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_Main 
   Caption         =   "BD Analyser"
   ClientHeight    =   7710
   ClientLeft      =   5400
   ClientTop       =   2085
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   3945
   Begin VB.CommandButton BO_Creer 
      Caption         =   "Crée Table"
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   6600
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CD2 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton BO_Sql 
      Caption         =   "BD -> SQL"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton BO_Htlm 
      Caption         =   "BD -> HTML"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox ZT_Sec 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0"
      Top             =   6240
      Width           =   495
   End
   Begin VB.TextBox ZT_Relations 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   7320
      Width           =   495
   End
   Begin VB.TextBox ZT_Requetes 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   6960
      Width           =   495
   End
   Begin VB.TextBox ZT_Tables 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      Top             =   6600
      Width           =   495
   End
   Begin MSComctlLib.TreeView TV_BD 
      Height          =   5415
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   9551
      _Version        =   393217
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton BO_Analyse 
      Caption         =   "Ouvrir une BD"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Ouverture d'un DataBase"
      Filter          =   "*.mdb"
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "sec"
      Height          =   195
      Left            =   2160
      TabIndex        =   12
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Temps Analyse:"
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   6240
      Width           =   1125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Relation(s):"
      Height          =   195
      Left            =   720
      TabIndex        =   9
      Top             =   7320
      Width           =   795
   End
   Begin VB.Label ET_Version 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2880
      TabIndex        =   7
      Top             =   480
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Requête(s):"
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   6960
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Table(s):"
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Structure de la BD Version:"
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   1920
   End
End
Attribute VB_Name = "F_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Jasmin Marinelli & Marie-Michelle Lavoie
'Analyser une base de données Acess, écrire le code sql pour crée les
'tables, insérer les enregistrements et construit un fichier html avec toutes les
'propriété de base de données
'8 Décembre 2000

Private Sub BO_Analyse_Click()
'On Error GoTo ExitTest
Dim BD As Database
Dim Table As TableDef
Dim RS As Recordset
Dim Champs As Field
Dim N As Node
Dim Requete As QueryDef
Dim err As Integer
err = 0
Dim Re As Relation
Dim MyTimer As Single

CD.FileName = ""
CD.InitDir = App.Path
CD.ShowOpen

MyTimer = Timer         'vérifie si le fichier est bien une base de données MDB

If CD.FileName = "" Or UCase(Right(CD.FileName, 3)) <> "MDB" Then GoTo ExitTest

Set BD = OpenDatabase(CD.FileName)  'ouverture de la base
TV_BD.Nodes.Clear
Set N = TV_BD.Nodes.Add(, , "BD", CD.FileTitle)

Set N = TV_BD.Nodes.Add("BD", tvwChild, "N_Tables", "Tables")


For Each Table In BD.TableDefs
    If Left(Table.Name, 4) <> "MSys" Then   'Pour chaque table, enlève les tables systèmes Msys
        Set N = TV_BD.Nodes.Add("N_Tables", tvwChild, Table.Name, Table.Name) 'inscrit le nom des tables
        Set RS = BD.OpenRecordset(Table.Name)
        
        For Each Champs In RS.Fields
            Set N = TV_BD.Nodes.Add(Table.Name, tvwChild, Table.Name + Champs.Name, Champs.Name)
        Next Champs
    Else
        err = err + 1
    End If
Next Table
    ZT_Tables = BD.TableDefs.Count - err    'compte le nombre de tables
Set N = TV_BD.Nodes.Add("BD", tvwChild, "N_Requetes", "Requêtes")
For Each Requete In BD.QueryDefs    'pour chaque requête, inscrit le nom des requêtes
    Set N = TV_BD.Nodes.Add("N_Requetes", tvwChild, Requete.Name, Requete.Name)
    Set N = TV_BD.Nodes.Add(Requete.Name, tvwChild, Requete.Name + "SQL", Requete.SQL)
Next Requete
    ZT_Requetes = BD.QueryDefs.Count    'compte le nombre de requêtes
    ET_Version = BD.Version             'inscrit la version de la base de données
    
    Set N = TV_BD.Nodes.Add("BD", tvwChild, "N_Relations", "Relations")
For Each Re In BD.Relations     'pour chaque relations, inscrit le nom des relations
    Set N = TV_BD.Nodes.Add("N_Relations", tvwChild, Re.Name, Re.Name)
Next Re
    ZT_Relations = BD.Relations.Count   'compte le nombre de requêtes
    ZT_Sec = Timer - MyTimer            'calcule le temps de l'exécution
    
ExitTest:
End Sub

Private Sub BO_Parcourir_Click()
CD.InitDir = App.Path       'ouvre la boîte dialogue "ouverture de fichiers"
CD.ShowOpen
End Sub

Private Sub BO_Creer_Click()
On Error GoTo ExitTest
Dim BD As Database
Dim TB As TableDef
Dim RS As Recordset
Dim F As Field
Dim MyTimer As Single
Dim NoChamps As Integer
Dim ChampsNom() As String
Dim i As Integer
CD2.DialogTitle = "Code de création de la table"
'CD2.Filter "*.Txt"
CD2.InitDir = App.Path
CD2.FileName = Left(CD.FileTitle, Len(CD.FileTitle) - 4) + "Table.txt"  'le nom de la base de données + Table.txt
CD2.ShowSave    'ouverture de la boîte de dialogue "Enregistrer"

Open CD2.FileName For Output As #1  'ouvrir en écriture le fichier html
MyTimer = Timer

'vérifie si c'est bien une base de données
If CD.FileName = "" Or UCase(Right(CD.FileName, 3)) <> "MDB" Then GoTo ExitTest
Set BD = OpenDatabase(CD.FileName)

Print #1, "Code pour la création de la table."; CD.FileTitle


For Each TB In BD.TableDefs     'pour chaque table dans la base de données
    If Left(TB.Name, 4) <> "MSys" Then      'enlève les tables systèmes
            
        Print #1, "CREATE TABLE "; TB.Name; " ("
        
        Set RS = BD.OpenRecordset(TB.Name)
        NoChamps = 0
        ReDim ChampsNom(1 To RS.Fields.Count)   'tableau des noms des champs
        
        For Each F In RS.Fields             'pour chaque champs
            NoChamps = NoChamps + 1         'compte le nombre de champs
            ChampsNom(NoChamps) = F.Name    'inscrit le nom du champ dans le tableau
        Next F
        
        For i = 1 To (NoChamps - 1)
        
'#################################################
'insérer le type et enlevé la liste sur la form
            Print #1, ChampsNom(i); ", "    'inscrit le nom du champ dans le fichier
        Next i

       Print #1, ChampsNom(i)
       Print #1, ");"
    End If
Next TB
Close #1        'ferme le fichier txt
MsgBox Timer - MyTimer, 0, "Information!"   'compte le temps d'analyse
ShellExecute hwnd, "open", CD2.FileName, "", "", vbNormalFocus  'ouverture du fichier
ExitTest:

End Sub

Private Sub BO_Htlm_Click()
On Error GoTo ExitTest
Dim BD As Database
Dim TB As TableDef
Dim RS As Recordset
Dim F As Field
Dim MyTimer As Single
Dim NoChamps As Integer
Dim ChampsNom() As String
Dim i As Integer
CD2.DialogTitle = "Sauver un HTML"
'CD2.Filter "*.html"
CD2.InitDir = App.Path
CD2.FileName = Left(CD.FileTitle, Len(CD.FileTitle) - 3) + "html" 'le nom de la base de données + html
CD2.ShowSave         'ouverture de la boîte de dialogue "Enregistrer=

Open CD2.FileName For Output As #1  'ouvrir en écriture le fichier html
MyTimer = Timer

'vérifie si c'est bien une base de données
If CD.FileName = "" Or UCase(Right(CD.FileName, 3)) <> "MDB" Then GoTo ExitTest
Set BD = OpenDatabase(CD.FileName)

Print #1, "<HTML>"      'inscrit les Tag dans le fichier html
Print #1, "<HEAD>"
Print #1, "<TITLE>"
Print #1, CD.FileTitle
Print #1, "</TITLE>"
Print #1, "</HEAD>"
Print #1, "<BODY>"

For Each TB In BD.TableDefs     'pour chaque table dans la base de données
    If Left(TB.Name, 4) <> "MSys" Then  'enlève les tables systèmes
        Print #1, "<H3>Table:</H3><H4>"; TB.Name; "</H4>"   'inscrit le nom de la table
        Set RS = BD.OpenRecordset(TB.Name)
        Print #1, "<table width=""75""border=""1"">"
        Print #1, "<tr>"
        NoChamps = 0
        ReDim ChampsNom(1 To RS.Fields.Count)   'tableau des noms des champs
        For Each F In RS.Fields                 'pour chaque champ
            NoChamps = NoChamps + 1             'compte le nombre de champs
            Print #1, "<td><B>"; F.Name; "</B></td>"    'inscrit le nom du champs dans le html
            ChampsNom(NoChamps) = F.Name                'inscrit le nom du champs dans le tableau
        Next F
        Print #1, "</tr>"
        Do Until RS.EOF
            Print #1, "<tr>"
            For i = 1 To UBound(ChampsNom)      'inscrit les enregistrements
                Print #1, "<td>"; RS.Fields(ChampsNom(i)); "</td>"
            Next i
            Print #1, "</tr>"
            RS.MoveNext
        Loop
        Print #1, "</table>"
        Print #1, "<HR>"
        Print #1, "<BR>"
    End If
Next TB

Print #1, "</BODY>"
Print #1, "</HEAD>"
Close #1            'ferme le html
MsgBox Timer - MyTimer, 0, "Information!"   'compte le nombre de temps d'analyse
ShellExecute hwnd, "open", CD2.FileName, "", "", vbNormalFocus  'ouverture du fichier
ExitTest:
End Sub

Private Sub BO_Sql_Click()
On Error GoTo ExitTest
Dim BD As Database
Dim TB As TableDef
Dim RS As Recordset
Dim F As Field
Dim MyTimer As Single
Dim NoChamps As Integer
Dim ChampsNom() As String
Dim NoEnregistrement As Integer
Dim Enregistrement() As String
Dim i As Integer
CD2.DialogTitle = "Sauver un TXT"
'CD2.Filter "*.Txt"
CD2.InitDir = App.Path
CD2.FileName = Left(CD.FileTitle, Len(CD.FileTitle) - 4) + "SQL.txt" 'le nom de la base de données + SQL.txt
CD2.ShowSave        'ouverture de la boîte de dialogue "Enregistrer"

Open CD2.FileName For Output As #1  'ouvre en écriture le fichier txt
MyTimer = Timer

'vérifie si c'est bien un base de données
If CD.FileName = "" Or UCase(Right(CD.FileName, 3)) <> "MDB" Then GoTo ExitTest
Set BD = OpenDatabase(CD.FileName)

Print #1, "Fichier du code SQL pour la BD"; CD.FileTitle


For Each TB In BD.TableDefs 'pour chaque table dans la base de données
    If Left(TB.Name, 4) <> "MSys" Then  'enlève les tables systems
        
        Print #1, "Table:   "; TB.Name  'inscrit le nom de la table
        
        Set RS = BD.OpenRecordset(TB.Name)
        NoChamps = 0
        NoEnregistrement = 0
        ReDim ChampsNom(1 To RS.Fields.Count)   'tableau des noms des champs
                    
        Print #1, "Insert Into "; TB.Name; " ( "
        
        For Each F In RS.Fields         'pour chaque champs
            NoChamps = NoChamps + 1     'compte le nombre de champ
            ChampsNom(NoChamps) = F.Name    'inscrit le nom du champ dans le tableau
        Next F
        
        For i = 1 To (NoChamps)
            Print #1, ChampsNom(i); ", "    'inscrit le nom du champ dans le txt
        Next i
        
        Print #1, " ) values ( "
        
        Do Until RS.EOF
           For i = 1 To (UBound(ChampsNom) - 1)
               Print #1, "'", RS.Fields(ChampsNom(i)); "'"; ", " 'inscrit les enregistrements
            Next i
            
            Print #1, "'", RS.Fields(ChampsNom(i)); "'"
            Print #1, ");"
            Print #1, "Insert Into "; TB.Name; " ( "
            RS.MoveNext
        Loop
        
       Print #1, ");"
    End If
Next TB
Close #1        'ferme le fichier txt
MsgBox Timer - MyTimer, 0, "Information!"   'compte le nombre de temps d'analyse
ShellExecute hwnd, "open", CD2.FileName, "", "", vbNormalFocus  'ouverture du fichier
ExitTest:
End Sub

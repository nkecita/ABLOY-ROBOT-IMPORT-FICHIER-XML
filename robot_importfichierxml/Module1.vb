
Imports System.Net
Imports System.Xml
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.Reflection
Imports System.Net.Mail

Public Structure configRobot
    Dim test As String
    Dim databasefichet As String
    Dim databasefichettest As String
    Dim filesxml As String
    Dim filesxmlbackup As String
    Dim filesxmlbackupbad As String
    Dim filestxtbackup As String
    Dim tables() As String
    Dim cylindres() As String
    Dim tabAdressePath() As String
    Dim tabClientsPath() As String
    Dim jourhisto As Integer
    Dim init As String

    Dim emailTo() As String
    Dim emailFrom As String
    Dim emailFromDisplay As String
    Dim emailSubjectOk As String
    Dim emailSubjectKo As String
    Dim emailBodyOk As String
    Dim emailBodyKo As String
    Dim smtpHost As String
    Dim smtpPort As String
    Dim smtpSLL As String
    Dim smtpLogin As String
    Dim smtpPassword As String
    Dim smtpCerfValidation As String
    Dim filescount As String


End Structure


Module Module1
    Public nbErreur As Integer = 0
    Public logError As String = String.Empty
    Public configGene As configRobot
    Public dsclient As New DataSet
    Public dsArticleXml As New DataSet
    Public dsArticle As New DataSet

    Sub Main()
        Try

            ' Lecture du fichier de configuration XML
            Dim startupPath As String
            Dim dsConfig As New DataSet
            startupPath = Assembly.GetExecutingAssembly().GetName().CodeBase
            startupPath = Path.GetDirectoryName(startupPath)
            dsConfig.ReadXml(startupPath + "\" + "config.xml")

            If dsConfig.Tables.Count = 1 Then
                With dsConfig.Tables(0).Rows(0)
                    configGene.test = .Item("TEST")
                    configGene.databasefichet = .Item("DATABASE-FICHET")
                    configGene.databasefichettest = .Item("DATABASE-FICHET-TEST")
                    configGene.filesxml = .Item("FILES-XML")
                    configGene.filesxmlbackup = .Item("FILES-XML-BACKUP")
                    configGene.filesxmlbackupbad = .Item("FILES-XML-BACKUP-BAD")
                    configGene.filestxtbackup = .Item("FILES-TXT-BACKUP")
                    configGene.tables = .Item("TABLES").ToString.Split(",")
                    configGene.cylindres = .Item("CYLINDRES").ToString.Split(",")
                    configGene.jourhisto = .Item("JOURHISTO")
                    configGene.init = .Item("INIT")
                    configGene.tabAdressePath = .Item("ADDRESS-FILE").ToString.Split(",")
                    configGene.tabClientsPath = .Item("CUSTTABLE-FILE").ToString.Split(",")
                    configGene.emailTo = .Item("EMAIL-TO").ToString.Split(",")
                    configGene.emailFrom = .Item("EMAIL-FROM").ToString()
                    configGene.emailFromDisplay = .Item("EMAIL-FROM-DISPLAY").ToString()
                    configGene.emailSubjectOk = .Item("EMAIL-SUBJECT-OK").ToString()
                    configGene.emailSubjectKo = .Item("EMAIL-SUBJECT-KO").ToString()
                    configGene.emailBodyOk = .Item("EMAIL-BODY-OK").ToString()
                    configGene.emailBodyKo = .Item("EMAIL-BODY-KO").ToString()
                    configGene.smtpHost = .Item("SMTPHOST").ToString()
                    configGene.smtpPort = .Item("SMTPPORT").ToString()
                    configGene.smtpSLL = .Item("SMTPSLL").ToString()
                    configGene.smtpLogin = .Item("STMPLOGIN").ToString()
                    configGene.smtpPassword = .Item("SMTPPASSWORD").ToString()
                    configGene.smtpCerfValidation = .Item("SMTPBYPASSCERTIFICATEVALIDATION").ToString()
                    configGene.filescount = .Item("FILESCOUNT").ToString()


                End With
            End If

            ' Si jour férie ou fermé alors le traitement ne fait rien
            If is_dayoff() Then
                Return
            End If

            ' Destruction des anciens répertoires pour éviter une surcharge disque inutile
            purge_repertoire(configGene.filesxmlbackup)
            purge_repertoire(configGene.filesxmlbackupbad)

            ' Merge Adresse
            mergeXml("Address", configGene.tabAdressePath)
            ' Merge CustTable
            mergeXml("CustTable", configGene.tabClientsPath)

            Dim di As New IO.DirectoryInfo(configGene.filesxml)
            Dim aryFi As IO.FileInfo() = di.GetFiles("*.xml")
            Dim fi As IO.FileInfo
            If aryFi.Length < Val(configGene.filescount) Then
                nbErreur = nbErreur + 1
                Report("Aucun fichier XML à traiter dans le répertoire ou il manque un fichier")
                logError &= "- Aucun fichier XML à traiter dans le répertoire ou il manque un fichier <br />"
                Return
            End If

            ' Sauvegarde de la table client pour les password
            Console.WriteLine("Sauvegarde de la table client pour la récupération des mots de passe")
            If sauvegarde_info(configGene.init) = False Then
                nbErreur = nbErreur + 1
                Report("Impossible de sauvegarder la table client pour récupération des mots de passe (fichier InventTableModule.XML introuvable)")
                logError &= "- Impossible de sauvegarder la table client pour récupération des mots de passe (fichier InventTableModule.XML introuvable)<br />"
                Return
            End If

            ' Modification des fichiers XML pour gérer les caractères spéciaux
            For Each fi In aryFi

                Dim streamRead As New IO.StreamReader(configGene.filesxml & "\" & fi.Name)
                Dim ContenuSourceXmlNew, ContenuSourceXmlOld As String
                ContenuSourceXmlNew = streamRead.ReadToEnd()
                streamRead.Close()
                ContenuSourceXmlOld = ContenuSourceXmlNew
                ContenuSourceXmlNew = Remplace_car_speciaux(ContenuSourceXmlNew)

                Dim cDir As String
                cDir = ""

                cDir = Date.Today.Year.ToString
                cDir = cDir & Date.Today.Month.ToString.PadLeft(2, "0")
                cDir = cDir & Date.Today.Day.ToString.PadLeft(2, "0")

                If Directory.Exists(configGene.filesxmlbackup & "\" & cDir) = False Then
                    Directory.CreateDirectory(configGene.filesxmlbackup & "\" & cDir)
                End If

                Dim streamWrite As New IO.StreamWriter(configGene.filesxmlbackup & "\" & cDir & "\" & fi.Name)
                streamWrite.WriteLine(ContenuSourceXmlOld)
                streamWrite.Close()
                fi.Delete()
            Next

            For f = 0 To configGene.tables.Length - 1
                If import_OK(configGene.tables(f).ToString) Then

                    ' Sauvegarde des données dans les fichiers TXT
                    Console.WriteLine("Sauvegarde des tables de la base de données en fichier TXT")
                    sauvegarde_TABLE(configGene.tables(f).ToString)

                    ' Purge des tables SQL
                    Console.WriteLine("Initialisation des tables de la base de données")
                    purge_TABLE(configGene.tables(f).ToString)
                    Try

                        ' Import des fichiers XML dans les tables SQL
                        Console.WriteLine("Importation des tables de la base de données")
                        import_table(configGene.tables(f).ToString)

                    Catch
                        Console.WriteLine(configGene.tables(f).ToString)
                        Console.WriteLine("Erreur dans le traitement : " & ErrorToString())
                        log(configGene.tables(f).ToString)
                        log(ErrorToString)
                        nbErreur = nbErreur + 1
                    End Try
                End If
            Next
        Catch
            Console.WriteLine("Erreur dans le traitement : " & ErrorToString())
            log(ErrorToString)
            nbErreur = nbErreur + 1
        Finally
            sendByMail(logError)
        End Try
        Report("")
    End Sub
    Function recup_info(ByVal wCode As String, ByVal wChamp As String) As String

        For i = 0 To dsclient.Tables(0).Rows.Count - 1
            If dsclient.Tables(0).Rows(i).Item("codeclient").ToString = wCode Then
                Return dsclient.Tables(0).Rows(i).Item(wChamp).ToString
            End If
        Next
        Return ""
    End Function
    Function recup_info_articleXml(ByVal wCode As String, ByVal wChamp As String) As String

        For i = 0 To dsArticleXml.Tables(0).Rows.Count - 1
            If dsArticleXml.Tables(0).Rows(i).Item("itemid").ToString = wCode Then
                Return dsArticleXml.Tables(0).Rows(i).Item(wChamp).ToString
            End If
        Next
        Return ""
    End Function
    Function recup_info_article(ByVal wCode As String, ByVal wChamp As String) As String

        For i = 0 To dsArticle.Tables(0).Rows.Count - 1
            If dsArticle.Tables(0).Rows(i).Item("num_article").ToString = wCode Then
                Return dsArticle.Tables(0).Rows(i).Item(wChamp).ToString
            End If
        Next
        Return ""
    End Function
    Function sauvegarde_info(ByVal winit As String) As Boolean
        Try
            Dim csql As String
            If winit = "OUI" Then
                csql = "SELECT codeclient,password,adv,blocage,langue FROM CLIENTS"
            Else
                csql = "SELECT codeclient,password,adv,blocageweb,langue FROM CLIENTS_NEW"
            End If


            Dim cCon As String
            cCon = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)

            Dim adapter = New OleDb.OleDbDataAdapter(csql, cCon)
            adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
            adapter.Fill(dsclient, "list_password")

            Dim cSql2 As String = "SELECT num_article,blocageWeb FROM ARTICLES_NEW"
            Dim adapter2 = New OleDb.OleDbDataAdapter(cSql2, cCon)
            adapter2.MissingSchemaAction = MissingSchemaAction.AddWithKey
            adapter2.Fill(dsArticle, "list_Article")

            If File.Exists(configGene.filesxml + "\" + "inventtableModule.xml") Then

                dsArticleXml.ReadXml(configGene.filesxml + "\" + "inventtableModule.xml")

            End If
        Catch
            Return False

        End Try
        Return True
    End Function
    Function import_table(ByVal wNameTable) As String

        Dim cdir As String

        cdir = Date.Today.Year.ToString
        cdir = cdir & Date.Today.Month.ToString.PadLeft(2, "0")
        cdir = cdir & Date.Today.Day.ToString.PadLeft(2, "0")


        Select Case wNameTable
            ' Traitement de la table client
            Case "CLIENTS_NEW"
                Console.WriteLine("-----> Importation de la table CLIENTS_NEW")
                import_client(cdir)
                ' Traitement de la table Adresse
            Case "ADRESSES_NEW"
                Console.WriteLine("-----> Importation de la table ADRESSES_NEW")
                import_adresse(cdir)
                ' Traitement de la table des codes postaux
            Case "CODESPOSTAUX"
                Console.WriteLine("-----> Importation de la table CODESPOSTAUX")
                import_zipcode(cdir)
                ' Traitement de la table des Articles
            Case "ARTICLES_NEW"
                Console.WriteLine("-----> Importation de la table des ARTICLES")
                import_article(cdir)
                ' Traitement de la table des libelles Articles
            Case "ARTICLES_LIBELLE"
                Console.WriteLine("-----> Importation de la table des LIBELLES ARTICLES")
                import_article_libelle(cdir)
                ' Traitement de la table des Tarifs Articles
            Case "ARTICLES_TARIF"
                Console.WriteLine("-----> Importation de la table des TARIFS ARTICLES")
                import_article_tarif(cdir)
                ' Traitement de la table des groupes de taxe client
            Case "GROUPE_TAXES_CLIENT"
                Console.WriteLine("-----> Importation de la table des GROUPES DE TAXE CLIENT")
                import_groupe_taxes_client(cdir)
                ' Traitement de la table des groupes de taxe Article
            Case "GROUPE_TAXES_ARTICLE"
                Console.WriteLine("-----> Importation de la table des GROUPES DE TAXE ARTICLE")
                import_groupe_taxes_article(cdir)
                ' Traitement de la table des groupes de taxe Valeur
            Case "GROUPE_TAXES_VALEUR"
                Console.WriteLine("-----> Importation de la table des GROUPES DE TAXE VALEUR")
                import_groupe_taxes_valeur(cdir)
                ' Traitement de la table Groupe Frais
            Case "GROUPE_FRAIS"
                Console.WriteLine("-----> Importation de la table des GROUPES DE FRAIS")
                import_groupe_frais(cdir)
            Case "CONDITIONS_LIVRAISON"
                Console.WriteLine("-----> Importation des conditions de livraison")
                import_ConditionLivraison(cdir)

        End Select
        Return ""
    End Function
    Function import_client(ByVal wDir As String) As String

        Try

            Dim dsTableXml As New DataSet
            dsTableXml.ReadXml(configGene.filesxmlbackup & "\" & wDir & "\custtable.xml")


            For i = 0 To dsTableXml.Tables(0).Rows.Count - 1

                Dim strRequete As String


                Dim connectionstring As String

                Dim oSqlDataAdapter As New OleDb.OleDbDataAdapter
                connectionstring = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)
                Dim cCon As New OleDb.OleDbConnection(connectionstring)



                strRequete = " INSERT INTO CLIENTS_NEW (codeclient,coderfr,raisonsociale,codepays,langue,email,codetarif,password,cleecon," & _
                             " adv,groupetaxe,CondPaiement, BlocageWeb,BlocageAx,GroupeFrais,condLiv,Bu) " & _
                             " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

                cCon.Open()
                oSqlDataAdapter.InsertCommand = New OleDb.OleDbCommand(strRequete, cCon)

                If is_column(dsTableXml, "AAF_Webid") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@codeclient", OleDb.OleDbType.VarChar, 8).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("AAF_Webid")), "", dsTableXml.Tables(0).Rows(i).Item("AAF_Webid").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@codeclient", OleDb.OleDbType.VarChar, 8).Value = ""
                End If

                If is_column(dsTableXml, "AccountNum") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@coderfr", OleDb.OleDbType.VarChar, 20).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("AccountNum")), "", dsTableXml.Tables(0).Rows(i).Item("AccountNum").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@coderfr", OleDb.OleDbType.VarChar, 20).Value = ""
                End If

                If is_column(dsTableXml, "Name") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@raisonsociale", OleDb.OleDbType.VarChar, 50).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("Name")), "", Remplace_car_speciaux(dsTableXml.Tables(0).Rows(i).Item("Name").ToString))
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@raisonsociale", OleDb.OleDbType.VarChar, 50).Value = ""
                End If

                If is_column(dsTableXml, "Country") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@codepays", OleDb.OleDbType.VarChar, 2).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("Country")), "", dsTableXml.Tables(0).Rows(i).Item("Country").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@codepays", OleDb.OleDbType.VarChar, 2).Value = ""
                End If


                Dim cLangue As String
                If is_column(dsTableXml, "LanguageId") Then

                    If IsDBNull(dsTableXml.Tables(0).Rows(i).Item("LanguageId")) Then
                        cLangue = ""
                    Else

                        Select Case Left(Trim(dsTableXml.Tables(0).Rows(i).Item("LanguageId").ToString), 2).ToUpper
                            Case "EN"
                                cLangue = "en-GB"
                            Case "FR"
                                cLangue = "fr-FR"
                            Case "IT"
                                cLangue = "it-IT"
                            Case "ES"
                                cLangue = "es-ES"
                            Case "PT"
                                cLangue = "pt-PT"
                            Case "NL"
                                cLangue = "nl-BE"
                            Case "SK"
                                cLangue = "sk-SK"


                            Case Else
                                If Trim(dsTableXml.Tables(0).Rows(i).Item("LanguageId").ToString).Length = 2 Then
                                    cLangue = Trim(dsTableXml.Tables(0).Rows(i).Item("LanguageId").ToString).ToLower & _
                                                       "-" & Trim(dsTableXml.Tables(0).Rows(i).Item("LanguageId").ToString).ToUpper
                                Else
                                    cLangue = dsTableXml.Tables(0).Rows(i).Item("LanguageId").ToString()
                                End If

                        End Select


                    End If
                Else
                    cLangue = ""
                End If



                'Etape 1 : Transtypage des langues connues

                'Fichier XML	Base SQL
                'Langue commençant par en	en-GB
                'Langue commençant par fr	fr-FR
                'Langue commençant par it	it-IT
                'Langue commençant par es	es-ES
                'Langue commençant par pt	pt-PT
                'Langue commençant par nl	nl-BE
                'Langue commençant par sk	sk-SK

                'Etape 2 : langue inconnue sans pays (2 caractères)
                'La langue xx devient xx-XX

                'Etape 3 : langue inconnue avec pays (5 caractères)
                'Copiée telle quelle


                oSqlDataAdapter.InsertCommand.Parameters.Add("@langue", OleDb.OleDbType.VarChar, 10).Value = cLangue

                'oSqlDataAdapter.InsertCommand.Parameters.Add("@langue", OleDb.OleDbType.VarChar, 10).Value = IIf(recup_info(dsTableXml.Tables(0).Rows(i).Item("AAF_Webid").ToString, "LANGUE") = "", cLangue, recup_info(dsTableXml.Tables(0).Rows(i).Item("AAF_Webid").ToString, "LANGUE"))

                If is_column(dsTableXml, "Email") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@email", OleDb.OleDbType.VarChar, 50).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("Email")), "", dsTableXml.Tables(0).Rows(i).Item("Email").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@email", OleDb.OleDbType.VarChar, 50).Value = ""
                End If

                If is_column(dsTableXml, "PriceGroup") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@codetarif", OleDb.OleDbType.VarChar, 50).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("PriceGroup")), "", dsTableXml.Tables(0).Rows(i).Item("PriceGroup").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@codetarif", OleDb.OleDbType.VarChar, 50).Value = ""
                End If

                oSqlDataAdapter.InsertCommand.Parameters.Add("@password", OleDb.OleDbType.VarChar, 50).Value = recup_info(dsTableXml.Tables(0).Rows(i).Item("AAF_Webid").ToString, "PASSWORD")


                oSqlDataAdapter.InsertCommand.Parameters.Add("@cleecon", OleDb.OleDbType.VarChar, 50).Value = ""



                oSqlDataAdapter.InsertCommand.Parameters.Add("@adv", OleDb.OleDbType.Boolean).Value = False 'IIf(recup_info(dsTableXml.Tables(0).Rows(i).Item("AAF_Webid").ToString, "ADV") = "0", False, True)

                If is_column(dsTableXml, "TaxGroup") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@groupetaxe", OleDb.OleDbType.VarChar, 10).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("TaxGroup")), "", dsTableXml.Tables(0).Rows(i).Item("TaxGroup").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@groupetaxe", OleDb.OleDbType.VarChar, 10).Value = ""
                End If

                If is_column(dsTableXml, "PaymTermid") Then

                    oSqlDataAdapter.InsertCommand.Parameters.Add("@CondPaiement", OleDb.OleDbType.Boolean).Value = IIf(dsTableXml.Tables(0).Rows(i).Item("PaymTermid").ToString.Trim = "00N", True, False)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@CondPaiement", OleDb.OleDbType.Boolean).Value = False
                End If

                If configGene.init = "OUI" Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@BlocageWeb", OleDb.OleDbType.Boolean).Value = IIf(recup_info(dsTableXml.Tables(0).Rows(i).Item("AAF_Webid").ToString, "BLOCAGE") = "0", False, True)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@BlocageWeb", OleDb.OleDbType.Boolean).Value = IIf(recup_info(dsTableXml.Tables(0).Rows(i).Item("AAF_Webid").ToString, "BLOCAGEWEB").Trim = "1", True, False)
                End If



                If is_column(dsTableXml, "Blocked") Then

                    oSqlDataAdapter.InsertCommand.Parameters.Add("@BlocageAx", OleDb.OleDbType.Boolean).Value = IIf(dsTableXml.Tables(0).Rows(i).Item("Blocked").ToString.Trim = "1" Or dsTableXml.Tables(0).Rows(i).Item("Blocked").ToString.Trim = "2", True, False)
                Else

                    oSqlDataAdapter.InsertCommand.Parameters.Add("@BlocageAx", OleDb.OleDbType.Boolean).Value = False
                End If

                If is_column(dsTableXml, "MarkupGroup") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@GroupeFrais", OleDb.OleDbType.VarChar, 20).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("MarkupGroup")), "", dsTableXml.Tables(0).Rows(i).Item("MarkupGroup").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@GroupeFrais", OleDb.OleDbType.VarChar, 20).Value = ""
                End If

                If is_column(dsTableXml, "DlvTerm") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@CondLiv", OleDb.OleDbType.VarChar, 6).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("DlvTerm")), "", dsTableXml.Tables(0).Rows(i).Item("DlvTerm").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@CondLiv", OleDb.OleDbType.VarChar, 6).Value = ""
                End If

                If is_column(dsTableXml, "BusinessUnit") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@Bu", OleDb.OleDbType.VarChar, 10).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("BusinessUnit")), "", dsTableXml.Tables(0).Rows(i).Item("BusinessUnit").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@Bu", OleDb.OleDbType.VarChar, 10).Value = ""
                End If
                oSqlDataAdapter.InsertCommand.ExecuteNonQuery()
                cCon.Close()
            Next

            Return String.Empty
        Catch ex As Exception
            Dim msg As String = "- CLIENTS_NEW = KO (Problèmes d'insertion dans la table)<br />"
            logError &= msg
            Console.WriteLine(msg)
            log(ex.ToString())
            Return String.Empty
        End Try
    End Function
    Function import_adresse(ByVal wDir As String) As String

        Try
            Dim dsTableXml As New DataSet
            dsTableXml.ReadXml(configGene.filesxmlbackup & "\" & wDir & "\address.xml")


            For i = 0 To dsTableXml.Tables(0).Rows.Count - 1

                Dim strRequete As String


                Dim connectionstring As String

                Dim oSqlDataAdapter As New OleDb.OleDbDataAdapter
                connectionstring = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)
                Dim cCon As New OleDb.OleDbConnection(connectionstring)



                strRequete = " INSERT INTO ADRESSES_NEW (codeclient,codeadresse,designation,adresse1,adresse2,adresse3," & _
                             " codepostal,localite,responsable,tel,fax,typliv,zone,CodeTournee,pays,PublieWeb) " & _
                             " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

                cCon.Open()
                oSqlDataAdapter.InsertCommand = New OleDb.OleDbCommand(strRequete, cCon)


                'Code Client
                If is_column(dsTableXml, "Addrrecid") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@codeclient", OleDb.OleDbType.VarChar, 8).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("Addrrecid")), "", dsTableXml.Tables(0).Rows(i).Item("Addrrecid").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@codeclient", OleDb.OleDbType.VarChar, 8).Value = ""
                End If


                ' Code Adresse
                If is_column(dsTableXml, "AAF_Webid") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@codeadresse", OleDb.OleDbType.VarChar, 6).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("AAF_Webid")), "", IIf(dsTableXml.Tables(0).Rows(i).Item("AAF_Webid").ToString.Trim = "0", "000000", dsTableXml.Tables(0).Rows(i).Item("AAF_Webid").ToString.Trim))
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@codeadresse", OleDb.OleDbType.VarChar, 6).Value = ""
                End If




                ' Raison Sociale
                If is_column(dsTableXml, "Name") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@designation", OleDb.OleDbType.VarChar, 50).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("Name")), "", Remplace_car_speciaux(dsTableXml.Tables(0).Rows(i).Item("Name").ToString))
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@designation", OleDb.OleDbType.VarChar, 50).Value = ""
                End If

                ' Ligne d'adresse
                Dim adresse() As String
                adresse = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("street")), "", Remplace_car_speciaux(dsTableXml.Tables(0).Rows(i).Item("street").ToString)).ToString.Trim.Split(vbLf)
                If is_column(dsTableXml, "street") Then

                    Select Case adresse.Length
                        Case 0
                            oSqlDataAdapter.InsertCommand.Parameters.Add("@adresse1", OleDb.OleDbType.VarChar, 50).Value = ""
                            oSqlDataAdapter.InsertCommand.Parameters.Add("@adresse2", OleDb.OleDbType.VarChar, 50).Value = ""
                            oSqlDataAdapter.InsertCommand.Parameters.Add("@adresse3", OleDb.OleDbType.VarChar, 50).Value = ""
                        Case 1
                            oSqlDataAdapter.InsertCommand.Parameters.Add("@adresse1", OleDb.OleDbType.VarChar, 50).Value = adresse(0)
                            oSqlDataAdapter.InsertCommand.Parameters.Add("@adresse2", OleDb.OleDbType.VarChar, 50).Value = ""
                            oSqlDataAdapter.InsertCommand.Parameters.Add("@adresse3", OleDb.OleDbType.VarChar, 50).Value = ""
                        Case 2
                            oSqlDataAdapter.InsertCommand.Parameters.Add("@adresse1", OleDb.OleDbType.VarChar, 50).Value = adresse(0)
                            oSqlDataAdapter.InsertCommand.Parameters.Add("@adresse2", OleDb.OleDbType.VarChar, 50).Value = adresse(1)
                            oSqlDataAdapter.InsertCommand.Parameters.Add("@adresse3", OleDb.OleDbType.VarChar, 50).Value = ""
                        Case 3
                            oSqlDataAdapter.InsertCommand.Parameters.Add("@adresse1", OleDb.OleDbType.VarChar, 50).Value = adresse(0)
                            oSqlDataAdapter.InsertCommand.Parameters.Add("@adresse2", OleDb.OleDbType.VarChar, 50).Value = adresse(1)
                            oSqlDataAdapter.InsertCommand.Parameters.Add("@adresse3", OleDb.OleDbType.VarChar, 50).Value = adresse(2)

                        Case Else
                            'Dim adresse2 As String
                            'adresse2 = ""
                            'For r = 1 To adresse.Length - 1
                            '    adresse2 = adresse2.Trim & IIf(adresse2 = "", "", " ") & adresse(r).Trim
                            'Next
                            'oSqlDataAdapter.InsertCommand.Parameters.Add("@adresse1", OleDb.OleDbType.VarChar, 50).Value = adresse(0)
                            'oSqlDataAdapter.InsertCommand.Parameters.Add("@adresse2", OleDb.OleDbType.VarChar, 50).Value = adresse2.Trim
                            oSqlDataAdapter.InsertCommand.Parameters.Add("@adresse1", OleDb.OleDbType.VarChar, 50).Value = adresse(0)
                            oSqlDataAdapter.InsertCommand.Parameters.Add("@adresse2", OleDb.OleDbType.VarChar, 50).Value = adresse(1)
                            oSqlDataAdapter.InsertCommand.Parameters.Add("@adresse3", OleDb.OleDbType.VarChar, 50).Value = adresse(2)

                    End Select


                End If

                'ZipCode
                If is_column(dsTableXml, "zipcode") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@CodePostal", OleDb.OleDbType.VarChar, 10).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("ZipCode")), "", dsTableXml.Tables(0).Rows(i).Item("ZipCode").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@CodePostal", OleDb.OleDbType.VarChar, 10).Value = ""
                End If

                ' City
                If is_column(dsTableXml, "City") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@localite", OleDb.OleDbType.VarChar, 50).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("city")), "", Remplace_car_speciaux(dsTableXml.Tables(0).Rows(i).Item("city").ToString))
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@localite", OleDb.OleDbType.VarChar, 50).Value = ""
                End If


                ' Telex
                If is_column(dsTableXml, "telex") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@responsable", OleDb.OleDbType.VarChar, 50).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("telex")), "", Remplace_car_speciaux(dsTableXml.Tables(0).Rows(i).Item("telex").ToString))
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@responsable", OleDb.OleDbType.VarChar, 50).Value = ""
                End If


                ' Phone
                If is_column(dsTableXml, "phone") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@tel", OleDb.OleDbType.VarChar, 30).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("phone")), "", dsTableXml.Tables(0).Rows(i).Item("phone").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@tel", OleDb.OleDbType.VarChar, 30).Value = ""
                End If

                ' Telefax
                If is_column(dsTableXml, "telefax") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@fax", OleDb.OleDbType.VarChar, 30).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("telefax")), "", dsTableXml.Tables(0).Rows(i).Item("telefax").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@fax", OleDb.OleDbType.VarChar, 30).Value = ""
                End If

                ' DivMode
                If is_column(dsTableXml, "DlvMode") Then

                    oSqlDataAdapter.InsertCommand.Parameters.Add("@typliv", OleDb.OleDbType.Boolean).Value = IIf(dsTableXml.Tables(0).Rows(i).Item("DlvMode").ToString.Trim = "70", True, False)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@typliv", OleDb.OleDbType.Boolean).Value = False
                End If

                ' FreightZone
                If is_column(dsTableXml, "FreightZone") Then

                    oSqlDataAdapter.InsertCommand.Parameters.Add("@Zone", OleDb.OleDbType.VarChar, 30).Value = If(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("FreightZone")), "", dsTableXml.Tables(0).Rows(i).Item("FreightZone").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@Zone", OleDb.OleDbType.VarChar, 30).Value = ""
                End If

                'Code Tournée
                If is_column(dsTableXml, "AAF_DestinationCodeId") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@CodeTournee", OleDb.OleDbType.VarChar, 2).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("AAF_DestinationCodeId")), "", dsTableXml.Tables(0).Rows(i).Item("AAF_DestinationCodeId").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@CodeTournee", OleDb.OleDbType.VarChar, 2).Value = ""
                End If

                ' Code Pays
                If is_column(dsTableXml, "Country") Then

                    oSqlDataAdapter.InsertCommand.Parameters.Add("@pays", OleDb.OleDbType.VarChar, 10).Value = If(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("Country")), "", dsTableXml.Tables(0).Rows(i).Item("Country").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@pays", OleDb.OleDbType.VarChar, 10).Value = ""
                End If

                ' Code Pays
                If is_column(dsTableXml, "PublieWeb") Then

                    oSqlDataAdapter.InsertCommand.Parameters.Add("@PublieWeb", OleDb.OleDbType.VarChar, 3).Value = If(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("PublieWeb")), "", dsTableXml.Tables(0).Rows(i).Item("PublieWeb").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@PublieWeb", OleDb.OleDbType.VarChar, 3).Value = ""
                End If

                oSqlDataAdapter.InsertCommand.ExecuteNonQuery()
                cCon.Close()
            Next
            Return String.Empty
        Catch ex As Exception
            Dim msg As String = "- ADRESSES_NEW = KO (Problèmes d'insertion dans la table)<br />"
            logError &= msg
            Console.WriteLine(msg)
            log(ex.ToString())
            Return String.Empty
        End Try
    End Function
    Function import_zipcode(ByVal wDir As String) As String

        Try
            Dim dsTableXml As New DataSet
            dsTableXml.ReadXml(configGene.filesxmlbackup & "\" & wDir & "\zipcode.xml")


            For i = 0 To dsTableXml.Tables(0).Rows.Count - 1

                Dim strRequete As String


                Dim connectionstring As String

                Dim oSqlDataAdapter As New OleDb.OleDbDataAdapter
                connectionstring = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)
                Dim cCon As New OleDb.OleDbConnection(connectionstring)



                strRequete = " INSERT INTO CodesPostaux (codepostal,ville,pays) VALUES (?,?,?)"

                cCon.Open()
                oSqlDataAdapter.InsertCommand = New OleDb.OleDbCommand(strRequete, cCon)

                ' Code Adresse
                If is_column(dsTableXml, "zipcode") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@codepostal", OleDb.OleDbType.VarChar, 15).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("zipcode")), "", dsTableXml.Tables(0).Rows(i).Item("zipcode").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@codepostal", OleDb.OleDbType.VarChar, 15).Value = ""
                End If

                'Code Client
                If is_column(dsTableXml, "city") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@ville", OleDb.OleDbType.VarChar, 50).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("city")), "", Remplace_car_speciaux(dsTableXml.Tables(0).Rows(i).Item("city").ToString))
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@ville", OleDb.OleDbType.VarChar, 50).Value = ""
                End If

                'Pays
                If is_column(dsTableXml, "country") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@pays", OleDb.OleDbType.VarChar, 50).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("country")), "", Remplace_car_speciaux(dsTableXml.Tables(0).Rows(i).Item("country").ToString))
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@pays", OleDb.OleDbType.VarChar, 50).Value = ""
                End If

                oSqlDataAdapter.InsertCommand.ExecuteNonQuery()
                cCon.Close()
            Next
            Return String.Empty
        Catch ex As Exception
            Dim msg As String = "- CodesPostaux = KO (Problèmes d'insertion dans la table)<br />"
            logError &= msg
            Console.WriteLine(msg)
            log(ex.ToString())
            Return String.Empty
        End Try
    End Function
    Function import_article(ByVal wDir As String) As String

        Try
            Dim dsTableXml As New DataSet
            dsTableXml.ReadXml(configGene.filesxmlbackup & "\" & wDir & "\inventtable.xml")

            For i = 0 To dsTableXml.Tables(0).Rows.Count - 1

                ' For i = 0 To dsTableXml.Tables(0).DefaultView.ToTable(True, "AAF_Webid").Rows.Count - 1
                Dim strRequete As String


                Dim connectionstring As String

                Dim oSqlDataAdapter As New OleDb.OleDbDataAdapter
                connectionstring = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)
                Dim cCon As New OleDb.OleDbConnection(connectionstring)

                Console.WriteLine(dsTableXml.Tables(0).Rows(i).Item("ItemId").ToString)

                strRequete = " INSERT INTO ARTICLES_NEW (num_article,num_axapta,code_tri,GroupeFrais,cond,origine,type_cde,cylindre,A2P,Groupetaxe,saisi,blocageweb,blocageAx) " & _
                             " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)"

                cCon.Open()
                oSqlDataAdapter.InsertCommand = New OleDb.OleDbCommand(strRequete, cCon)

                ' Code Article
                If is_column(dsTableXml, "AAF_Webid") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@num_article", OleDb.OleDbType.VarChar, 53).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("AAF_Webid")), "", dsTableXml.Tables(0).Rows(i).Item("AAF_Webid").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@num_article", OleDb.OleDbType.VarChar, 53).Value = ""
                End If

                ' Code Article AXAPTA
                If is_column(dsTableXml, "ItemId") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@num_axapta", OleDb.OleDbType.VarChar, 53).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("ItemId")), "", dsTableXml.Tables(0).Rows(i).Item("ItemId").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@num_axapta", OleDb.OleDbType.VarChar, 53).Value = ""
                End If


                ' Code Tri
                If is_column(dsTableXml, "AAF_FamilleVenteAAW") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@code_tri", OleDb.OleDbType.VarChar, 53).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("AAF_FamilleVenteAAW")), "", dsTableXml.Tables(0).Rows(i).Item("AAF_FamilleVenteAAW").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@code_tri", OleDb.OleDbType.VarChar, 53).Value = ""
                End If


                ' Code Groupe de frais Article
                If is_column(dsTableXml, "ItemId") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@GroupeFrais", OleDb.OleDbType.VarChar, 20).Value = recup_info_articleXml(dsTableXml.Tables(0).Rows(i).Item("ItemId").ToString, "MarkupGroupId").ToString
                    'IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("MarkupGroupId")), "", dsTableXml.Tables(0).Rows(i).Item("MarkupGroupId").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@GroupeFrais", OleDb.OleDbType.VarChar, 20).Value = ""
                End If


                ' Conditionnement
                If is_column(dsTableXml, "AAF_QuantiteConditionnement") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@cond", OleDb.OleDbType.Integer).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("AAF_QuantiteConditionnement")), 0, dsTableXml.Tables(0).Rows(i).Item("AAF_QuantiteConditionnement"))
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@cond", OleDb.OleDbType.Integer).Value = 0
                End If

                ' Fournisseur Préférentiel
                If is_column(dsTableXml, "PrimaryVendorId") Then
                    If dsTableXml.Tables(0).Rows(i).Item("aaf_groupwebid").ToString.Trim.ToUpper <> "STD" Then
                        oSqlDataAdapter.InsertCommand.Parameters.Add("@origine", OleDb.OleDbType.VarChar, 10).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("PrimaryVendorId")), "", dsTableXml.Tables(0).Rows(i).Item("PrimaryVendorId").ToString)
                    Else
                        oSqlDataAdapter.InsertCommand.Parameters.Add("@origine", OleDb.OleDbType.VarChar, 10).Value = ""
                    End If

                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@origine", OleDb.OleDbType.VarChar, 10).Value = ""
                End If

                ' Type de commande
                If is_column(dsTableXml, "aaf_groupwebid") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@type_cde", OleDb.OleDbType.VarChar, 5).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("aaf_groupwebid")), "", dsTableXml.Tables(0).Rows(i).Item("aaf_groupwebid").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@type_cde", OleDb.OleDbType.VarChar, 5).Value = ""
                End If

                ' Type de cylinbdre
                If is_column(dsTableXml, "AAF_InventPlanNumberId") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@Cylindre", OleDb.OleDbType.VarChar, 15).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("AAF_InventPlanNumberId")), "", IIf(verif_plan(dsTableXml.Tables(0).Rows(i).Item("AAF_InventPlanNumberId").ToString.Trim) = True, dsTableXml.Tables(0).Rows(i).Item("AAF_InventPlanNumberId").ToString, ""))
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@Cylindre", OleDb.OleDbType.VarChar, 15).Value = ""
                End If

                ' Produit A2P ?0
                If is_column(dsTableXml, "aaf_a2p") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@A2p", OleDb.OleDbType.Boolean).Value = IIf(dsTableXml.Tables(0).Rows(i).Item("aaf_a2p").ToString.Trim = "1", True, False)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@A2p", OleDb.OleDbType.Boolean).Value = False
                End If

                If is_column(dsTableXml, "ItemId") Then
                    ' Code Taxe Article
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@GroupeTaxe", OleDb.OleDbType.VarChar, 20).Value = recup_info_articleXml(dsTableXml.Tables(0).Rows(i).Item("ItemId").ToString, "taxitemgroupid").ToString

                    ' Peut etre saisi en article standard
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@saisi", OleDb.OleDbType.Boolean).Value = IIf(recup_info_articleXml(dsTableXml.Tables(0).Rows(i).Item("ItemId").ToString, "AAF_interditVente").ToString = "0", False, True)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@GroupeTaxe", OleDb.OleDbType.VarChar, 20).Value = ""
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@saisi", OleDb.OleDbType.Boolean).Value = False
                End If

                ' Blocage Web ?
                If recup_info_article(dsTableXml.Tables(0).Rows(i).Item("ItemId").ToString, "blocageweb").ToString.ToLower = "true" Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@blocageweb", OleDb.OleDbType.Boolean).Value = True
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@blocageweb", OleDb.OleDbType.Boolean).Value = False
                End If

                ' Blocage Ax ?
                If is_column(dsTableXml, "ItemId") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@blocageAx", OleDb.OleDbType.Boolean).Value = IIf(recup_info_articleXml(dsTableXml.Tables(0).Rows(i).Item("ItemId").ToString, "Blocked").ToString = "0", False, True)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@blocageAx", OleDb.OleDbType.Boolean).Value = False
                End If
                oSqlDataAdapter.InsertCommand.ExecuteNonQuery()
                cCon.Close()
            Next
            Return String.Empty
        Catch ex As Exception
            Dim msg As String = "- ARTICLES_NEW = KO (Problèmes d'insertion dans la table)<br />"
            logError &= msg
            Console.WriteLine(msg)
            log(ex.ToString())
            Return String.Empty
        End Try
    End Function
    Function import_article_libelle(ByVal wDir As String) As String

        Try
            Dim dsTableXml As New DataSet
            dsTableXml.ReadXml(configGene.filesxmlbackup & "\" & wDir & "\inventTxt.xml")


            For i = 0 To dsTableXml.Tables(0).Rows.Count - 1

                Dim strRequete As String


                Dim connectionstring As String

                Dim oSqlDataAdapter As New OleDb.OleDbDataAdapter
                connectionstring = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)
                Dim cCon As New OleDb.OleDbConnection(connectionstring)



                strRequete = " INSERT INTO ARTICLES_LIBELLE (num_article,langue,libelle,bu) " & _
                             " VALUES (?,?,?,?)"

                cCon.Open()
                oSqlDataAdapter.InsertCommand = New OleDb.OleDbCommand(strRequete, cCon)

                ' Code Article
                If is_column(dsTableXml, "ItemId") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@num_article", OleDb.OleDbType.VarChar, 53).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("ItemId")), "", dsTableXml.Tables(0).Rows(i).Item("ItemId").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@num_article", OleDb.OleDbType.VarChar, 53).Value = ""
                End If

                ' Langue
                Dim cLangue As String
                If is_column(dsTableXml, "LanguageId") Then

                    If IsDBNull(dsTableXml.Tables(0).Rows(i).Item("LanguageId")) Then
                        cLangue = ""
                    Else

                        Select Case Left(Trim(dsTableXml.Tables(0).Rows(i).Item("LanguageId").ToString), 2).ToUpper
                            Case "EN"
                                cLangue = "en-GB"
                            Case "FR"
                                cLangue = "fr-FR"
                            Case "IT"
                                cLangue = "it-IT"
                            Case "ES"
                                cLangue = "es-ES"
                            Case "PT"
                                cLangue = "pt-PT"
                            Case "NL"
                                cLangue = "nl-BE"
                            Case "SK"
                                cLangue = "sk-SK"


                            Case Else
                                If Trim(dsTableXml.Tables(0).Rows(i).Item("LanguageId").ToString).Length = 2 Then
                                    cLangue = Trim(dsTableXml.Tables(0).Rows(i).Item("LanguageId").ToString).ToLower & _
                                                       "-" & Trim(dsTableXml.Tables(0).Rows(i).Item("LanguageId").ToString).ToUpper
                                Else
                                    cLangue = dsTableXml.Tables(0).Rows(i).Item("LanguageId").ToString()
                                End If

                        End Select


                    End If
                Else
                    cLangue = ""
                End If



                'Etape 1 : Transtypage des langues connues

                'Fichier XML	Base SQL
                'Langue commençant par en	en-GB
                'Langue commençant par fr	fr-FR
                'Langue commençant par it	it-IT
                'Langue commençant par es	es-ES
                'Langue commençant par pt	pt-PT
                'Langue commençant par nl	nl-BE
                'Langue commençant par sk	sk-SK

                'Etape 2 : langue inconnue sans pays (2 caractères)
                'La langue xx devient xx-XX

                'Etape 3 : langue inconnue avec pays (5 caractères)
                'Copiée telle quelle



                oSqlDataAdapter.InsertCommand.Parameters.Add("@langue", OleDb.OleDbType.VarChar, 10).Value = cLangue

                'If is_column(dsTableXml, "languageId") Then
                'Select Case Len(dsTableXml.Tables(0).Rows(i).Item("languageId"))
                '   Case 2
                'oSqlDataAdapter.InsertCommand.Parameters.Add("@langue", OleDb.OleDbType.VarChar, 10).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("languageId")), "", dsTableXml.Tables(0).Rows(i).Item("languageId").ToString.Trim.ToLower & "-" & dsTableXml.Tables(0).Rows(i).Item("languageId").ToString.Trim.ToUpper)

                '   Case 5
                'Dim tmp As String
                'tmp = Left(dsTableXml.Tables(0).Rows(i).Item("languageId").ToString.Trim, 2).Trim.ToLower
                'tmp = tmp & "-"
                'tmp = Right(dsTableXml.Tables(0).Rows(i).Item("languageId").ToString.Trim, 2).Trim.ToUpper

                '                oSqlDataAdapter.InsertCommand.Parameters.Add("@langue", OleDb.OleDbType.VarChar, 10).Value = tmp

                '                   Case Else
                '              oSqlDataAdapter.InsertCommand.Parameters.Add("@langue", OleDb.OleDbType.VarChar, 10).Value = ""
                '             End Select

                '            Else
                '           oSqlDataAdapter.InsertCommand.Parameters.Add("@langue", OleDb.OleDbType.VarChar, 10).Value = ""
                '          End If

                ' Libelle
                If is_column(dsTableXml, "txt") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@libelle", OleDb.OleDbType.VarChar, 255).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("txt")), "", Remplace_car_speciaux(dsTableXml.Tables(0).Rows(i).Item("txt")).ToString.Replace("'", "''"))

                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@libelle", OleDb.OleDbType.VarChar, 255).Value = ""
                End If

                ' bu
                If is_column(dsTableXml, "BusinessUnit") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@bu", OleDb.OleDbType.VarChar, 25).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("BusinessUnit")), "", Remplace_car_speciaux(dsTableXml.Tables(0).Rows(i).Item("txt")).ToString.Replace("'", "''"))

                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@bu", OleDb.OleDbType.VarChar, 25).Value = ""
                End If

                oSqlDataAdapter.InsertCommand.ExecuteNonQuery()
                cCon.Close()
            Next

            Return String.Empty
        Catch ex As Exception
            Dim msg As String = "- ARTICLES_LIBELLE = KO (Problèmes d'insertion dans la table)<br />"
            logError &= msg
            Console.WriteLine(msg)
            log(ex.ToString())
            Return String.Empty
        End Try
    End Function
    Function import_article_tarif(ByVal wDir As String) As String
        Try
            Dim dsTableXml As New DataSet
            dsTableXml.ReadXml(configGene.filesxmlbackup & "\" & wDir & "\priceDiscTable.xml")


            For i = 0 To dsTableXml.Tables(0).Rows.Count - 1
                Console.WriteLine(IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("Itemrelation")), "", dsTableXml.Tables(0).Rows(i).Item("Itemrelation").ToString))
                Dim strRequete As String


                Dim connectionstring As String

                Dim oSqlDataAdapter As New OleDb.OleDbDataAdapter
                connectionstring = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)
                Dim cCon As New OleDb.OleDbConnection(connectionstring)



                strRequete = " INSERT INTO ARTICLES_TARIF (code_tarif,num_article,p_vente,qtemini,c_edition,relation) " & _
                             " VALUES (?,?,?,?,?,?)"

                cCon.Open()
                oSqlDataAdapter.InsertCommand = New OleDb.OleDbCommand(strRequete, cCon)

                ' Code Tarif
                If is_column(dsTableXml, "AccountRelation") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@code_tarif", OleDb.OleDbType.VarChar, 53).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("AccountRelation")), "", dsTableXml.Tables(0).Rows(i).Item("AccountRelation").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@code_tarif", OleDb.OleDbType.VarChar, 53).Value = ""
                End If

                ' Code Article
                If is_column(dsTableXml, "Itemrelation") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@num_article", OleDb.OleDbType.VarChar, 53).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("Itemrelation")), "", dsTableXml.Tables(0).Rows(i).Item("Itemrelation").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@num_article", OleDb.OleDbType.VarChar, 53).Value = ""
                End If

                ' Prix
                If is_column(dsTableXml, "Amount") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@p_vente", OleDb.OleDbType.Numeric).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("Amount")), "", dsTableXml.Tables(0).Rows(i).Item("Amount").ToString.Replace(".", ","))
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@p_vente", OleDb.OleDbType.Numeric).Value = ""
                End If

                ' QteMini
                If is_column(dsTableXml, "QuantityAmount") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@qtemini", OleDb.OleDbType.Integer).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("QuantityAmount")), "", dsTableXml.Tables(0).Rows(i).Item("QuantityAmount").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@qtemini", OleDb.OleDbType.Integer).Value = ""
                End If

                ' Code Edition
                If is_column(dsTableXml, "AAF_EditCode") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@c_edition", OleDb.OleDbType.VarChar, 10).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("AAF_EditCode")), "", dsTableXml.Tables(0).Rows(i).Item("AAF_EditCode").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@c_edition", OleDb.OleDbType.VarChar, 10).Value = ""
                End If

                ' Relation
                If is_column(dsTableXml, "relation") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@relation", OleDb.OleDbType.VarChar, 1).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("relation")), "", dsTableXml.Tables(0).Rows(i).Item("relation").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@relation", OleDb.OleDbType.VarChar, 1).Value = ""
                End If


                oSqlDataAdapter.InsertCommand.ExecuteNonQuery()
                cCon.Close()
            Next
            Return String.Empty
        Catch ex As Exception
            Dim msg As String = "- ARTICLES_TARIF = KO (Problèmes d'insertion dans la table)<br />"
            logError &= msg
            Console.WriteLine(msg)
            log(ex.ToString())

            Return String.Empty
        End Try
    End Function
    Function import_groupe_taxes_client(ByVal wDir As String) As String
        Try
            Dim dsTableXml As New DataSet
            dsTableXml.ReadXml(configGene.filesxmlbackup & "\" & wDir & "\TaxGroupData.xml")


            For i = 0 To dsTableXml.Tables(0).Rows.Count - 1

                Dim strRequete As String


                Dim connectionstring As String

                Dim oSqlDataAdapter As New OleDb.OleDbDataAdapter
                connectionstring = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)
                Dim cCon As New OleDb.OleDbConnection(connectionstring)



                strRequete = " INSERT INTO GROUPE_TAXES_CLIENT (groupe,code) " & _
                             " VALUES (?,?)"

                cCon.Open()
                oSqlDataAdapter.InsertCommand = New OleDb.OleDbCommand(strRequete, cCon)

                ' Groupe de Taxe Client
                If is_column(dsTableXml, "TaxGroup") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@groupe", OleDb.OleDbType.VarChar, 20).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("TaxGroup")), "", dsTableXml.Tables(0).Rows(i).Item("TaxGroup").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@groupe", OleDb.OleDbType.VarChar, 20).Value = ""
                End If

                ' Code Taxe
                If is_column(dsTableXml, "TaxCode") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@code", OleDb.OleDbType.VarChar, 20).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("TaxCode")), "", dsTableXml.Tables(0).Rows(i).Item("TaxCode").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@code", OleDb.OleDbType.VarChar, 20).Value = ""
                End If

                oSqlDataAdapter.InsertCommand.ExecuteNonQuery()
                cCon.Close()
            Next
            Return String.Empty
        Catch ex As Exception
            Dim msg As String = "- GROUPE_TAXES_CLIENT = KO (Problèmes d'insertion dans la table)<br />"
            logError &= msg
            Console.WriteLine(msg)
            log(ex.ToString())
            Return String.Empty
        End Try
    End Function
    Function import_groupe_taxes_article(ByVal wDir As String) As String
        Try
            Dim dsTableXml As New DataSet
            dsTableXml.ReadXml(configGene.filesxmlbackup & "\" & wDir & "\TaxOnItem.xml")


            For i = 0 To dsTableXml.Tables(0).Rows.Count - 1

                Dim strRequete As String


                Dim connectionstring As String

                Dim oSqlDataAdapter As New OleDb.OleDbDataAdapter
                connectionstring = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)
                Dim cCon As New OleDb.OleDbConnection(connectionstring)



                strRequete = " INSERT INTO GROUPE_TAXES_ARTICLE (groupe,code) " & _
                             " VALUES (?,?)"

                cCon.Open()
                oSqlDataAdapter.InsertCommand = New OleDb.OleDbCommand(strRequete, cCon)

                ' Groupe de Taxe Artticle
                If is_column(dsTableXml, "TaxItemGroup") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@groupe", OleDb.OleDbType.VarChar, 20).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("TaxItemGroup")), "", dsTableXml.Tables(0).Rows(i).Item("TaxItemGroup").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@groupe", OleDb.OleDbType.VarChar, 20).Value = ""
                End If

                ' Code Taxe
                If is_column(dsTableXml, "TaxCode") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@code", OleDb.OleDbType.VarChar, 20).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("TaxCode")), "", dsTableXml.Tables(0).Rows(i).Item("TaxCode").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@code", OleDb.OleDbType.VarChar, 20).Value = ""
                End If

                oSqlDataAdapter.InsertCommand.ExecuteNonQuery()
                cCon.Close()
            Next
            Return String.Empty
        Catch ex As Exception
            Dim msg As String = "- GROUPE_TAXES_ARTICLE = KO (Problèmes d'insertion dans la table)<br />"
            logError &= msg
            Console.WriteLine(msg)
            log(ex.ToString())
            Return String.Empty
        End Try
    End Function
    Function import_groupe_taxes_valeur(ByVal wDir As String) As String
        Try
            Dim dsTableXml As New DataSet
            dsTableXml.ReadXml(configGene.filesxmlbackup & "\" & wDir & "\TaxData.xml")


            For i = 0 To dsTableXml.Tables(0).Rows.Count - 1

                Dim strRequete As String


                Dim connectionstring As String

                Dim oSqlDataAdapter As New OleDb.OleDbDataAdapter
                connectionstring = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)
                Dim cCon As New OleDb.OleDbConnection(connectionstring)



                strRequete = " INSERT INTO GROUPE_TAXES_VALEUR (code,taux) " & _
                             " VALUES (?,?)"

                cCon.Open()
                oSqlDataAdapter.InsertCommand = New OleDb.OleDbCommand(strRequete, cCon)

                ' Groupe de Taxe 
                If is_column(dsTableXml, "taxCode") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@code", OleDb.OleDbType.VarChar, 20).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("taxCode")), "", dsTableXml.Tables(0).Rows(i).Item("taxCode").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@code", OleDb.OleDbType.VarChar, 20).Value = ""
                End If

                ' Taux à Appliquer
                If is_column(dsTableXml, "taxValue") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@taux", OleDb.OleDbType.Numeric).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("taxValue")), 0, CDec(dsTableXml.Tables(0).Rows(i).Item("taxValue").ToString.Replace(".", ",")))
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@taux", OleDb.OleDbType.Numeric).Value = ""
                End If

                oSqlDataAdapter.InsertCommand.ExecuteNonQuery()
                cCon.Close()
            Next
            Return String.Empty
        Catch ex As Exception
            Dim msg As String = "- GROUPE_TAXES_VALEUR = KO (Problèmes d'insertion dans la table)<br />"
            logError &= msg
            Console.WriteLine(msg)
            log(ex.ToString())
            Return String.Empty
        End Try
    End Function
    Function import_groupe_frais(ByVal wDir As String) As String
        Try
            Dim dsTableXml As New DataSet
            dsTableXml.ReadXml(configGene.filesxmlbackup & "\" & wDir & "\MarkupAutoTable.xml")


            For i = 0 To dsTableXml.Tables(0).Rows.Count - 1

                Dim strRequete As String


                Dim connectionstring As String

                Dim oSqlDataAdapter As New OleDb.OleDbDataAdapter
                connectionstring = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)
                Dim cCon As New OleDb.OleDbConnection(connectionstring)



                strRequete = " INSERT INTO GROUPE_FRAIS (groupeFraisClient,groupeFraisArticle,type,code,valeur) " & _
                             " VALUES (?,?,?,?,?)"

                cCon.Open()
                oSqlDataAdapter.InsertCommand = New OleDb.OleDbCommand(strRequete, cCon)

                ' Clé vers le Client
                If is_column(dsTableXml, "AccountRelation") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@GroupeFraisClient", OleDb.OleDbType.VarChar, 20).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("AccountRelation")), "", dsTableXml.Tables(0).Rows(i).Item("AccountRelation").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@GroupeFraisClient", OleDb.OleDbType.VarChar, 20).Value = ""
                End If

                ' Clé vers L'article
                If is_column(dsTableXml, "ItemRelation") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@GroupeFraisArticle", OleDb.OleDbType.VarChar, 20).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("ItemRelation")), "", dsTableXml.Tables(0).Rows(i).Item("ItemRelation").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@GroupeFraisArticle", OleDb.OleDbType.VarChar, 20).Value = ""
                End If



                ' Type de frais
                If is_column(dsTableXml, "MarkupCategory") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@type", OleDb.OleDbType.VarChar, 20).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("MarkupCategory")), "", dsTableXml.Tables(0).Rows(i).Item("MarkupCategory").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@type", OleDb.OleDbType.VarChar, 20).Value = ""
                End If

                ' Code Frais
                If is_column(dsTableXml, "MarkupCode") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@code", OleDb.OleDbType.VarChar, 20).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("MarkupCode")), "", dsTableXml.Tables(0).Rows(i).Item("MarkupCode").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@code", OleDb.OleDbType.VarChar, 20).Value = ""
                End If

                ' Taux à Appliquer
                If is_column(dsTableXml, "Value") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@valeur", OleDb.OleDbType.Numeric).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("Value")), 0, CDec(dsTableXml.Tables(0).Rows(i).Item("Value").ToString.Replace(".", ",")))
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@valeur", OleDb.OleDbType.Numeric).Value = 0
                End If

                oSqlDataAdapter.InsertCommand.ExecuteNonQuery()
                cCon.Close()
            Next
            Return String.Empty
        Catch ex As Exception
            Dim msg As String = "- GROUPE_FRAIS = KO (Problèmes d'insertion dans la table)<br />"
            logError &= msg
            Console.WriteLine(msg)
            log(ex.ToString())
            Return String.Empty
        End Try
    End Function
    Function import_ConditionLivraison(ByVal wDir As String) As String
        Try
            Dim dsTableXml As New DataSet
            dsTableXml.ReadXml(configGene.filesxmlbackup & "\" & wDir & "\DlvTerm.xml")


            For i = 0 To dsTableXml.Tables(0).Rows.Count - 1

                Dim strRequete As String


                Dim connectionstring As String

                Dim oSqlDataAdapter As New OleDb.OleDbDataAdapter
                connectionstring = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)
                Dim cCon As New OleDb.OleDbConnection(connectionstring)



                strRequete = " INSERT INTO CONDITIONS_LIVRAISON (code,codefrais,limite,valeur) " & _
                             " VALUES (?,?,?,?)"

                cCon.Open()
                oSqlDataAdapter.InsertCommand = New OleDb.OleDbCommand(strRequete, cCon)

                ' Code condition
                If is_column(dsTableXml, "Code") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@code", OleDb.OleDbType.VarChar, 10).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("Code")), "", dsTableXml.Tables(0).Rows(i).Item("Code").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@code", OleDb.OleDbType.VarChar, 10).Value = ""
                End If

                ' Code frais
                If is_column(dsTableXml, "AAF_FrancoMarkupCode") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@CodeFrais", OleDb.OleDbType.VarChar, 10).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("AAF_FrancoMarkupCode")), "", dsTableXml.Tables(0).Rows(i).Item("AAF_FrancoMarkupCode").ToString)
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@CodeFrais", OleDb.OleDbType.VarChar, 10).Value = ""
                End If



                ' Limite
                If is_column(dsTableXml, "AAF_FrancoAmount") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@limite", OleDb.OleDbType.Numeric).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("AAF_FrancoAmount")), 0, CDec(dsTableXml.Tables(0).Rows(i).Item("AAF_FrancoAmount").ToString.Replace(".", ",")))
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@limite", OleDb.OleDbType.Numeric).Value = 0
                End If


                ' Valeur
                If is_column(dsTableXml, "AAF_FrancoFixedAmount") Then
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@valeur", OleDb.OleDbType.Numeric).Value = IIf(IsDBNull(dsTableXml.Tables(0).Rows(i).Item("AAF_FrancoFixedAmount")), 0, CDec(dsTableXml.Tables(0).Rows(i).Item("AAF_FrancoFixedAmount").ToString.Replace(".", ",")))
                Else
                    oSqlDataAdapter.InsertCommand.Parameters.Add("@valeur", OleDb.OleDbType.Numeric).Value = 0
                End If

                oSqlDataAdapter.InsertCommand.ExecuteNonQuery()
                cCon.Close()
            Next
            Return String.Empty
        Catch ex As Exception
            Dim msg As String = "- CONDITIONS_LIVRAISON = KO (Problèmes d'insertion dans la table)<br />"
            logError &= msg
            Console.WriteLine(msg)
            log(ex.ToString())
            Return String.Empty
        End Try
    End Function
    Function is_column(ByVal wDs As DataSet, ByVal wname As String) As Boolean
        For i = 0 To wDs.Tables(0).Columns.Count - 1
            If wDs.Tables(0).Columns(i).ColumnName.ToUpper = wname.ToUpper Then
                Return True

            End If
        Next

        Return False
    End Function
    Function purge_TABLE(ByVal wNameTable) As String
        Try

            Dim cmdDelete As New OleDb.OleDbCommand
            Dim conSql As New OleDb.OleDbConnection
            Dim cCon As String
            cCon = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)

            conSql.ConnectionString = cCon
            conSql.Open()
            cmdDelete.Connection = conSql


            Select Case wNameTable
                Case "CLIENTS_NEW"
                    cmdDelete.CommandText = "DELETE FROM  " & wNameTable & " WHERE LEFT(codeclient,1)<>'E'"
                Case "ADRESSES_NEW"
                    cmdDelete.CommandText = "DELETE FROM  " & wNameTable
                Case Else
                    cmdDelete.CommandText = "DELETE FROM  " & wNameTable
            End Select

            'If wNameTable.ToUpper = "CLIENTS_NEW" Then
            '    cmdDelete.CommandText = "DELETE FROM  " & wNameTable & " WHERE (LEFT(codeclient,1)<>'E' AND LEFT(codeclient,2)<>'80')"
            'Else
            '    cmdDelete.CommandText = "DELETE FROM  " & wNameTable
            'End If

            'If wNameTable.ToUpper = "ADRESSES_NEW" Then
            '    cmdDelete.CommandText = "DELETE FROM  " & wNameTable & " WHERE  LEFT(codeclient,2)<>'80'"
            'Else
            '    cmdDelete.CommandText = "DELETE FROM  " & wNameTable
            'End If



            cmdDelete.ExecuteNonQuery()
            conSql.Close()

            Return String.Empty
        Catch ex As Exception
            Dim msg As String = "- " & wNameTable & " = KO (Probleme d'initialisation de la table dans la base de données)<br />"
            logError &= msg
            Console.WriteLine(msg)
            log(ex.ToString())
            Return String.Empty
        End Try

    End Function
    Function sauvegarde_TABLE(ByVal wNameTable) As String
        Try

            Dim dsTable As DataSet = New DataSet()
            Dim cSql As String

            cSql = "SELECT * FROM " & wNameTable

            Dim cCon As String
            cCon = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)

            Dim adapter = New OleDb.OleDbDataAdapter(cSql, cCon)
            adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
            adapter.Fill(dsTable, "list_table")

            Dim cDir As String
            cDir = ""
            If dsTable.Tables.Count > 0 Then

                cDir = Date.Today.Year.ToString
                cDir = cDir & Date.Today.Month.ToString.PadLeft(2, "0")
                cDir = cDir & Date.Today.Day.ToString.PadLeft(2, "0")

                If Directory.Exists(configGene.filestxtbackup & "\" & cDir) = False Then
                    Directory.CreateDirectory(configGene.filestxtbackup & "\" & cDir)
                End If
            End If

            Dim cBackFile As String = configGene.filestxtbackup & "\" & cDir & "\" & wNameTable & ".txt"
            Dim streamWrite As New IO.StreamWriter(cBackFile)

            For i = 0 To dsTable.Tables(0).Rows.Count - 1
                Dim cLigne As String
                cLigne = ""
                For c = 0 To dsTable.Tables(0).Columns.Count - 1

                    cLigne = cLigne & dsTable.Tables(0).Rows(i).Item(c).ToString & vbTab
                Next
                streamWrite.WriteLine(cLigne)
            Next

            streamWrite.Close()
            Return String.Empty
        Catch ex As Exception
            Dim msg As String = "- " & wNameTable & " = KO (Problème de sauvegarde au format .TXT de la table)<br />"
            logError &= msg
            Console.WriteLine(msg)
            log(ex.ToString())
            Return String.Empty
        End Try
    End Function
    Function Remplace_car_speciaux(ByVal wContenu As String) As String

        Dim newcontenu As String = ""
        Do While newcontenu.Trim <> wContenu.Trim

            newcontenu = wContenu
            wContenu = wContenu.Replace("&lt;", "<")
            wContenu = wContenu.Replace("&gt;", ">")
            wContenu = wContenu.Replace("&apos;", "'")
            wContenu = wContenu.Replace("&quot;", """")

          
            wContenu = wContenu.Replace("&#115;", "é")
            wContenu = wContenu.Replace("&#138;", "Š")
            wContenu = wContenu.Replace("&#142;", "Ž")
            wContenu = wContenu.Replace("&#154;", "š")
            wContenu = wContenu.Replace("&#158;", "ž")
            wContenu = wContenu.Replace("&#32;", "é")
            wContenu = wContenu.Replace("&#99;", "é")
            wContenu = wContenu.Replace("&#186;", "º")
            wContenu = wContenu.Replace("&#170;", "ª")

            wContenu = wContenu.Replace("&#192;", "À")
            wContenu = wContenu.Replace("&#193;", "Á")
            wContenu = wContenu.Replace("&#194;", "Â")
            wContenu = wContenu.Replace("&#195;", "Ã")
            wContenu = wContenu.Replace("&#196;", "Ä")
            wContenu = wContenu.Replace("&#197;", "Å")
            wContenu = wContenu.Replace("&#198;", "Æ")
            wContenu = wContenu.Replace("&#199;", "Ç")
            wContenu = wContenu.Replace("&#200;", "È")
            wContenu = wContenu.Replace("&#201;", "É")
            wContenu = wContenu.Replace("&#202;", "Ê")
            wContenu = wContenu.Replace("&#203;", "Ë")
            wContenu = wContenu.Replace("&#204;", "Ì")
            wContenu = wContenu.Replace("&#205;", "Í")
            wContenu = wContenu.Replace("&#206;", "Î")
            wContenu = wContenu.Replace("&#207;", "Ï")
            wContenu = wContenu.Replace("&#208;", "Ð")
            wContenu = wContenu.Replace("&#209;", "Ñ")
            wContenu = wContenu.Replace("&#210;", "Ò")
            wContenu = wContenu.Replace("&#211;", "Ó")
            wContenu = wContenu.Replace("&#212;", "Ô")
            wContenu = wContenu.Replace("&#213;", "Õ")
            wContenu = wContenu.Replace("&#214;", "Ö")
            wContenu = wContenu.Replace("&#215;", "×")
            wContenu = wContenu.Replace("&#216;", "Ø")
            wContenu = wContenu.Replace("&#217;", "Ù")
            wContenu = wContenu.Replace("&#218;", "Ú")
            wContenu = wContenu.Replace("&#219;", "Û")
            wContenu = wContenu.Replace("&#220;", "Ü")
            wContenu = wContenu.Replace("&#221;", "Ý")
            wContenu = wContenu.Replace("&#222;", "Þ")
            wContenu = wContenu.Replace("&#223;", "ß")
            wContenu = wContenu.Replace("&#224;", "à")
            wContenu = wContenu.Replace("&#225;", "á")
            wContenu = wContenu.Replace("&#226;", "â")
            wContenu = wContenu.Replace("&#227;", "ã")
            wContenu = wContenu.Replace("&#228;", "ä")
            wContenu = wContenu.Replace("&#229;", "å")
            wContenu = wContenu.Replace("&#230;", "æ")
            wContenu = wContenu.Replace("&#231;", "ç")
            wContenu = wContenu.Replace("&#232;", "è")
            wContenu = wContenu.Replace("&#233;", "é")
            wContenu = wContenu.Replace("&#234;", "ê")
            wContenu = wContenu.Replace("&#235;", "ë")
            wContenu = wContenu.Replace("&#236;", "ì")
            wContenu = wContenu.Replace("&#237;", "í")
            wContenu = wContenu.Replace("&#238;", "î")
            wContenu = wContenu.Replace("&#239;", "ï")
            wContenu = wContenu.Replace("&#240;", "ð")
            wContenu = wContenu.Replace("&#241;", "ñ")
            wContenu = wContenu.Replace("&#242;", "ò")
            wContenu = wContenu.Replace("&#243;", "ó")
            wContenu = wContenu.Replace("&#244;", "ô")
            wContenu = wContenu.Replace("&#245;", "õ")
            wContenu = wContenu.Replace("&#246;", "ö")
            wContenu = wContenu.Replace("&#247;", "÷")
            wContenu = wContenu.Replace("&#248;", "ø")
            wContenu = wContenu.Replace("&#249;", "ù")
            wContenu = wContenu.Replace("&#250;", "ú")
            wContenu = wContenu.Replace("&#251;", "û")
            wContenu = wContenu.Replace("&#252;", "ü")
            wContenu = wContenu.Replace("&#253;", "ý")
            wContenu = wContenu.Replace("&#254;", "þ")
            wContenu = wContenu.Replace("&#255;", "ÿ")
            wContenu = wContenu.Replace("&#268;", "Č")


            wContenu = wContenu.Replace("&#270;", "Ď")
            wContenu = wContenu.Replace("&#327;", "Ň")
            wContenu = wContenu.Replace("&#352;", "Š")
            wContenu = wContenu.Replace("&#381;", "Ž")
            wContenu = wContenu.Replace("&#269;", "č")

            wContenu = wContenu.Replace("&#271;", "ď")
            wContenu = wContenu.Replace("&#314;", "Ĺ")
            wContenu = wContenu.Replace("&#317;", "Ľ")
            wContenu = wContenu.Replace("&#318;", "ľ")
            wContenu = wContenu.Replace("&#328;", "ň")
            wContenu = wContenu.Replace("&#340;", "Ŕ")
            wContenu = wContenu.Replace("&#341;", "ŕ")
            wContenu = wContenu.Replace("&#353;", "š")
            wContenu = wContenu.Replace("&#356;", "Ť")
            wContenu = wContenu.Replace("&#357;", "ť")
            wContenu = wContenu.Replace("&#382;", "ž")



            wContenu = wContenu.Replace("&#199;", "Ç")

            wContenu = wContenu.Replace("&sup2;", "²")
            wContenu = wContenu.Replace("&sup3;", "³")
            wContenu = wContenu.Replace("&#10;", vbCrLf)
            wContenu = wContenu.Replace("&amp;", "")



            If newcontenu.Trim = wContenu.Trim Then
                Exit Do
            End If
        Loop

        Return wContenu
    End Function
    Function is_dayoff() As Boolean
        Try
            Dim dsDayOff As DataSet = New DataSet()
            Dim csql As String

            csql = "SELECT dt_ferie FROM jours_feries WHERE dt_ferie = convert(varchar,getdate(),112)"

            Dim cCon As String
            cCon = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)

            Dim adapter = New OleDb.OleDbDataAdapter(csql, cCon)
            adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
            adapter.Fill(dsDayOff, "list_dayoff")

            If dsDayOff.Tables(0).Rows.Count = 1 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Dim msg As String = "- jours_feries = KO (Problèmes lié à la détection des jours feriés)<br />"
            logError &= msg
            Console.WriteLine(msg)
            log(ex.ToString())
            Return False
        End Try
    End Function
    Function purge_repertoire(ByVal pPath As String) As String
        Try
            Dim subdirectoryEntries As String()

            For i = 1 To 2
                If i = 1 Then
                    subdirectoryEntries = Directory.GetDirectories(pPath)
                Else
                    subdirectoryEntries = Directory.GetDirectories(pPath)
                End If

                Dim subdirectory As String
                For Each subdirectory In subdirectoryEntries

                    Dim cDateRep, cDirSup As String
                    cDateRep = Right(subdirectory, 8)
                    cDirSup = cDateRep.Substring(6, 2) & "/"
                    cDirSup = cDirSup & cDateRep.Substring(4, 2) & "/"
                    cDirSup = cDirSup & cDateRep.Substring(0, 4)

                    If DateDiff(DateInterval.Day, CDate(cDirSup), Today) > configGene.jourhisto Then
                        Directory.Delete(subdirectory, True)
                    End If

                Next subdirectory

            Next
            Return String.Empty
        Catch ex As Exception
            Dim msg As String = "- Problèmes de purge du répertoire '" & pPath & "')<br />"
            logError &= msg
            Console.WriteLine(msg)
            log(ex.ToString())
            Return String.Empty
        End Try
    End Function
    Function log(ByVal wErreur As String)

        Dim streamWrite As New IO.StreamWriter("log.txt", True)
        streamWrite.WriteLine("Date de Traitement : " & Date.Today.Day.ToString.PadLeft(2, "0") & "/" & Date.Today.Month.ToString.PadLeft(2, "0") & "/" & Date.Today.Year.ToString & " " & Now.Hour.ToString.PadLeft(2, "0") & ":" & Now.Minute.ToString.PadLeft(2, "0"))
        streamWrite.WriteLine("************************************************************************")
        streamWrite.WriteLine(wErreur)
        streamWrite.WriteLine("************************************************************************")

        streamWrite.Close()
    End Function
    Function Report(ByVal Msg As String)

        Dim streamWrite As New IO.StreamWriter("archive.txt", True)
        streamWrite.WriteLine("Date de Traitement : " & Date.Today.Day.ToString.PadLeft(2, "0") & "/" & Date.Today.Month.ToString.PadLeft(2, "0") & "/" & Date.Today.Year.ToString & " " & Now.Hour.ToString.PadLeft(2, "0") & ":" & Now.Minute.ToString.PadLeft(2, "0"))

        streamWrite.WriteLine(IIf(nbErreur = 0, "Traitement OK", "Le traitement comporte " & nbErreur.ToString & " erreur(s)"))
        If nbErreur > 0 Then
            If Msg <> "" Then
                streamWrite.WriteLine(Msg)
            End If

            streamWrite.WriteLine("Veillez consulter le fichier log.txt pour plus de détails")
        End If

        streamWrite.WriteLine("************************************************************************")



        streamWrite.Close()
    End Function
    Function import_OK(ByVal wNameTable) As Boolean
        Dim cdir As String

        cdir = Date.Today.Year.ToString
        cdir = cdir & Date.Today.Month.ToString.PadLeft(2, "0")
        cdir = cdir & Date.Today.Day.ToString.PadLeft(2, "0")

        Select Case wNameTable
            ' Traitement de la table client
            Case "CLIENTS_NEW"
                Return testFilesXml(configGene.filesxmlbackup & "\" & cdir & "\CustTable.xml", 50, "CustTable")
            Case "ADRESSES_NEW"
                Return testFilesXml(configGene.filesxmlbackup & "\" & cdir & "\Address.xml", 50, "Address")
            Case "CODESPOSTAUX"
                Return testFilesXml(configGene.filesxmlbackup & "\" & cdir & "\ZipCode.xml", 0, "ZipCode")
            Case "ARTICLES_NEW"
                Return testFilesXml(configGene.filesxmlbackup & "\" & cdir & "\inventTable.xml", 0, "inventTable")
            Case "ARTICLES_LIBELLE"
                Return testFilesXml(configGene.filesxmlbackup & "\" & cdir & "\InventTxt.xml", 0, "InventTxt")
            Case "ARTICLES_TARIF"
                Return testFilesXml(configGene.filesxmlbackup & "\" & cdir & "\PriceDiscTable.xml", 0, "PriceDiscTable")
            Case "GROUPE_TAXES_CLIENT"
                Return testFilesXml(configGene.filesxmlbackup & "\" & cdir & "\TaxGroupData.xml", 0, "TaxGroupData")
            Case "GROUPE_TAXES_ARTICLE"
                Return testFilesXml(configGene.filesxmlbackup & "\" & cdir & "\TaxOnItem.xml", 0, "TaxOnItem")
            Case "GROUPE_TAXES_VALEUR"
                Return testFilesXml(configGene.filesxmlbackup & "\" & cdir & "\TaxData.xml", 0, "TaxData")
            Case "GROUPE_FRAIS"
                Return testFilesXml(configGene.filesxmlbackup & "\" & cdir & "\MarkupAutoTable.xml", 0, "MarkupAutoTable")
            Case "CONDITIONS_LIVRAISON"
                Return testFilesXml(configGene.filesxmlbackup & "\" & cdir & "\DlvTerm.xml", 0, "DlvTerm")
        End Select
        Return True
    End Function
    Function verif_plan(ByVal wPlan As String) As Boolean

        Dim isHere As Boolean = configGene.cylindres.contains(wPlan)
        If isHere = True Then
            Return True
        End If

        Return False
    End Function
    Function mergeXml(ByVal pFile As String, ByVal pTable() As String) As Boolean
        Dim vPath As String = String.Empty
        Try
            Console.WriteLine("Merge de : " & pFile & ".xml")
            Dim xmlAdress As XmlDocument = New XmlDocument()
            xmlAdress.LoadXml("<?xml version=""1.0"" encoding=""iso-8859-1""?><" & pFile & "></" & pFile & ">")

            For Each Path As String In pTable
                vPath = Path
                If IO.File.Exists(configGene.filesxml & "\" & Path & ".xml") Then
                    Dim objXml As XmlDocument = New XmlDocument()
                    objXml.Load(configGene.filesxml & "\" & vPath & ".xml")

                    For Each objNode As XmlNode In objXml.DocumentElement.ChildNodes
                        Dim newNode As XmlNode = xmlAdress.ImportNode(objNode, True)
                        xmlAdress.DocumentElement.AppendChild(newNode)
                    Next
                    deplaceFile(configGene.filesxml & "\" & vPath & ".xml", configGene.filesxmlbackup)
                    xmlAdress.Save(configGene.filesxml & "\" & pFile & ".xml")
                Else
                    Dim msg As String = "- " & vPath & ".xml = KO (fichier introuvable)<br />"
                    logError &= msg
                    Console.WriteLine(msg)
                End If
            Next
            Console.WriteLine("Fin de Merge : " & pFile & ".xml")
            Return True
        Catch ex As Exception
            Dim msg As String = "- " & vPath & ".xml = KO (fichier XML erroné)<br />"
            logError &= msg
            Console.WriteLine(msg)
            deplaceFile(configGene.filesxml & "\" & vPath & ".xml", configGene.filesxmlbackupbad)
            log(ex.ToString())
            Return False
        End Try
    End Function
    Function deplaceFile(ByVal pFile As String, ByVal pDir As String) As Boolean
        Try

            Dim cDir As String
            cDir = ""

            cDir = Date.Today.Year.ToString
            cDir = cDir & Date.Today.Month.ToString.PadLeft(2, "0")
            cDir = cDir & Date.Today.Day.ToString.PadLeft(2, "0")

            If Directory.Exists(pDir & "\" & cDir) = False Then
                Directory.CreateDirectory(pDir & "\" & cDir)
            End If

            If File.Exists(pFile) Then
                File.Copy(pFile, pDir & "\" & cDir & "\" & Path.GetFileName(pFile), True)
                File.Delete(pFile)
            End If

        Catch ex As Exception
            log(ex.ToString())
            Return False
        End Try
    End Function
    Function testFilesXml(ByVal pFile As String, ByVal plimit As Integer, ByVal pFileName As String) As Boolean
        Try

            If IO.File.Exists(pFile) Then
                Dim dsTableXml As New DataSet
                dsTableXml.ReadXml(pFile)

                If dsTableXml.Tables(0).Rows.Count <= plimit Then

                    Dim msg As String = "- " & pFileName & ".xml = KO (Le fichier est incomplet < " & plimit & " enregistrements)<br />"
                    logError &= msg
                    Console.WriteLine(msg)
                    deplaceFile(pFile, configGene.filesxmlbackupbad)
                    Return False
                Else
                    Return True
                End If
            Else
                Dim msg As String = "- " & pFileName & ".xml = KO (fichier introuvable)<br />"
                logError &= msg
                Console.WriteLine(msg)
                Return False
            End If
        Catch ex As Exception
            Dim msg As String = "- " & pFileName & ".xml = KO (fichier erroné)<br />"
            logError &= msg
            Console.WriteLine(msg)
            deplaceFile(pFile, configGene.filesxmlbackupbad)
            log(ex.ToString())
            Return False
        End Try
    End Function
    Public Function sendByMail(ByVal pBody As String) As Boolean

        Try

            Dim oMail As New MailMessage()
            oMail.From = New MailAddress(configGene.emailFrom, configGene.emailFromDisplay)
            For Each pTo As String In configGene.emailTo
                oMail.To.Add(New MailAddress(pTo))
            Next

            oMail.IsBodyHtml = True

            If String.IsNullOrEmpty(pBody) Then
                oMail.Subject = configGene.emailSubjectOk
                oMail.Body = configGene.emailBodyOk
            Else
                oMail.Priority = MailPriority.High
                oMail.Subject = configGene.emailSubjectKo
                oMail.Body = String.Format(configGene.emailBodyKo, pBody)
            End If

            'oMail.Attachments.Add(New Attachment())

            Dim oSmtp As New SmtpClient()
            oSmtp.Host = configGene.smtpHost
            oSmtp.Port = Integer.Parse(configGene.smtpPort)
            oSmtp.Credentials = New NetworkCredential(configGene.smtpLogin, configGene.smtpPassword)
            oSmtp.EnableSsl = Boolean.Parse(configGene.smtpSLL)

            If Boolean.Parse(configGene.smtpCerfValidation) Then

            End If

            oSmtp.Send(oMail)
            Return True
        Catch ex As Exception
            log(ex.ToString())
            Return False
        End Try
    End Function


    'Public Function sendByMail(ByVal pBody As String) As Boolean
    '    Try
    '        Dim myMail As New System.Web.Mail.MailMessage()
    '        myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserver", configGene.smtpHost)
    '        myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserverport", configGene.smtpPort)
    '        myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusing", "2")
    '        myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1")
    '        myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", configGene.smtpLogin)
    '        myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", configGene.smtpPassword)
    '        myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpusessl", configGene.smtpSLL)
    '        myMail.From = configGene.emailFrom
    '        myMail.To = "elharraq91@gmail.com" 'Replace(configGene.emailTo, ",", ";")
    '        myMail.Subject = configGene.emailSubject
    '        myMail.BodyFormat = Web.Mail.MailFormat.Html
    '        myMail.Priority = System.Web.Mail.MailPriority.High

    '        If String.IsNullOrEmpty(pBody) Then
    '            myMail.Body = configGene.emailBodyOk
    '        Else
    '            myMail.Body = String.Format(configGene.emailBodyKo, pBody)
    '        End If

    '        System.Web.Mail.SmtpMail.SmtpServer = configGene.smtpHost & ":" & configGene.smtpPort
    '        System.Web.Mail.SmtpMail.Send(myMail)
    '        Return True
    '    Catch ex As Exception
    '        log(ex.ToString)
    '    End Try
    'End Function

End Module


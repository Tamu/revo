﻿Imports System.Xml
Imports System.Xml.XPath
Imports System.IO

''' -----------------------------------------------------------------------------
''' Project	 : GabXMLClass
''' Class	 : GabXML
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' GabXML permet une gestion facile des documents XML grâce à XPath.
''' </summary>
''' <remarks>
''' GabXML doit être instanciée avant d'être utilisée.
''' Voir par exemple http://www.mangue.org/wiki/XPath pour la syntaxe des requêtes XPath
''' </remarks>
''' <history>
''' 	[Gabriel]	11/05/2005	Created
'''     [Gabriel]   06/06/2009  Updated :
'''                 - Ajout des méhodes SelectValue, SelectFirstNode, SelectNodes et SetNode
'''                 - Suppression des méthodes en double : le paramètre index est maintenant optionnel (valeur par défaut : 1)
'''                 - Meilleure utilisation des requêtes XPath
'''                 - Position devient Index
'''                 - Une erreur est jetée si le fichier n'existe pas
'''                 - Compilé avec succès sous Visual Studio 2008
''' </history>
''' -----------------------------------------------------------------------------
Public Class RevoXML

    'Le nom du fichier xml sur lequel on travaillera
    Private fichier As String

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructeur de la classe
    ''' </summary>
    ''' <param name="leFichier"></param>
    ''' <remarks>
    ''' leFichier doit obligatoirement être un fichier existant
    ''' </remarks>
    ''' <history>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal leFichier As String)
        If System.IO.File.Exists(leFichier) Or Left(leFichier, 4).ToLower = "http" Then
            fichier = leFichier
        Else
            ' Throw New FileNotFoundException("Le fichier " & leFichier & " est introuvable !")
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructeur de la classe
    ''' </summary>
    ''' <remarks>
    ''' Ne nécessite pas de fichier xml existant
    ''' </remarks>
    ''' <history>
    ''' 	[Gabriel]	10/05/2005	Created
    '''     [Gabriel]   06/06/2009  Updated
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        fichier = ""
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' définit/retourne le nom du fichier
    ''' </summary>
    ''' <value>Un fichier .xml existant</value>
    ''' <remarks>
    ''' Veillez à indiquer un fichier XML existant!
    ''' </remarks>
    ''' <history>
    ''' 	[Gabriel]	10/05/2005	Created
    '''     [Gabriel]   06/06/2009  Updated
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property leFichier() As String
        Get
            Return fichier
        End Get
        Set(ByVal Value As String)
            fichier = Value
        End Set
    End Property

    ''' <summary>
    ''' Retourne la valeur du premier noeud correspondant à la requête XPath.
    ''' Cette fonction devrait être la plus utilisée car elle permet de retourner le résultat de n'importe quelle requête XPath.
    ''' Quelques exemples qui montrent la puissance de XPath :
    ''' /cd/piste[5]/titre/text() renvoie le titre de la piste 5
    ''' /cd/piste[@numero='3']/titre renvoie le titre de la piste dont l'attribut numéro est '3'
    ''' /cd[artiste='Pink Floyd'][3]/piste[5]/titre renvoie le titre de la 5ème piste du 3ème CD ayant "Pink Floyd" comme artiste.
    ''' /cd[artiste='Pink Floyd' and id='PF004']/piste[2]/duree renvoie la durée de la seconde piste du CD de Pink Floyd ayant l'ID 'PF004'
    ''' </summary>
    ''' <param name="query">Une requête XPath de sélection</param>
    ''' <returns>Une chaine de caractère contenant la valeur du noeud</returns>
    ''' <remarks></remarks>
    ''' <history>
    '''     [Gabriel]   06/06/2009  Created
    ''' </history>
    Public Function SelectValue(ByVal query As String) As String

        'déclarations
        Dim valeur As String = String.Empty
        Try
            'on charge le fichier xml
            Dim doc As New XmlDocument
            doc.Load(fichier)

            'execute et récupère la valeur de la requête
            valeur = doc.SelectSingleNode(query).InnerText

            'libère les ressources
            doc = Nothing
            
        Catch ex As Exception
            'en cas d'erreur
            'MsgBox(ex.Message)
            valeur = "ERRORXML"
        End Try

        Return valeur

    End Function

    ''' <summary>
    ''' Remplace le premier noeud correspondant à la requête XPath par un nouveau.
    ''' </summary>
    ''' <param name="xpath">Une requête XPath de sélection</param>
    ''' <param name="node">Un noeud (XmlNode)</param>
    ''' <remarks></remarks>
    ''' <history>
    '''     [Gabriel]   06/06/2009  Created
    ''' </history>
    Public Sub SetNode(ByVal xpath As String, ByVal node As XmlNode, Optional ByVal index As Integer = 1)
        Try

            'charge le fichier xml
            Dim doc As New XmlDocument
            doc.Load(fichier)

            'definit la racine de l'element a modifier
            Dim parent As XmlNode = doc.SelectNodes(xpath).ItemOf(index - 1).ParentNode

            'cherche le noeud à remplacer
            Dim n As XmlNode = doc.SelectNodes(xpath).ItemOf(index - 1)

            'on remplace par le nouveau noeud
            parent.ReplaceChild(n, node)

            'Sauve les modifications
            doc.Save(fichier)

            doc = Nothing

        Catch 'e As Exception

            MsgBox("setSettings <> Erreur dans la modification de " & fichier) '& " : " '& e.Message)

        End Try

    End Sub

    ''' <summary>
    ''' Retourne le premier noeud correspondant à la requête XPath, Nothing si aucun noeud n'est trouvé
    ''' </summary>
    ''' <param name="xpath">La requête XPath</param>
    ''' <returns>Un noeud (XmlNode)</returns>
    ''' <remarks></remarks>
    ''' <history>
    '''     [Gabriel]   06/06/2009  Created
    ''' </history>
    Public Function SelectFirstNode(ByVal xpath As String) As XmlNode

        'déclarations
        Dim n As XmlNode
        Try
            'on charge le fichier xml
            Dim doc As New XmlDocument
            doc.Load(fichier)

            'execute et récupère la valeur de la requête
            n = doc.SelectSingleNode(xpath)

            'libère les ressources
            doc = Nothing

        Catch ex As Exception
            'en cas d'erreur
            MsgBox(ex.Message)
            n = Nothing
        End Try

        Return n

    End Function

    ''' <summary>
    ''' Retourne la liste des noeuds correspondants à la requête XPath ou Nothing si rien n'est trouvé
    ''' </summary>
    ''' <param name="xpath">Une requête XPath</param>
    ''' <returns>Une liste de noeuds (XmlNodeList)</returns>
    ''' <remarks></remarks>
    ''' <history>
    '''     [Gabriel]   06/06/2009  Created
    ''' </history>
    Public Function SelectNodes(ByVal xpath As String) As XmlNodeList

        'déclarations
        Dim nl As XmlNodeList
        Try
            'on charge le fichier xml
            Dim doc As New XmlDocument
            doc.Load(fichier)

            'execute et récupère la valeur de la requête
            nl = doc.SelectNodes(xpath)

            'libère les ressources
            doc = Nothing

        Catch ex As Exception
            'en cas d'erreur
            'MsgBox(ex.Message)
            nl = Nothing
        End Try

        Return nl

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Objectif : Renvoie la valeur du n-ème élément pointé par la requête XPath
    ''' Attention : Le premier noeud possède l'index 1
    ''' </summary>
    ''' <param name="xpath">Une requête XPath</param>
    ''' <param name="index">L'index de l'élément</param>
    ''' <returns>Chaine de caractères</returns>
    ''' <remarks>
    ''' Utilisation |
    ''' exemple 1 : getElementValue("/polygone/point", 2) renvoie la valeur du 2ème "point" de "polygone" |
    ''' exemple 2 : getElementValue("/polygone/point[2]") renvoie la même chose
    ''' </remarks>
    ''' <history>
    ''' 	[Gabriel]	10/05/2005	Created
    '''     [Gabriel]   06/06/2009  Updated
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function getElementValue(ByVal xpath As String, Optional ByVal index As Integer = 1) As String

        'declarations
        Dim valeur As String = String.Empty
        'on charge le fichier xml

        Try
            Dim doc As New XmlDocument
            doc.Load(fichier)

            'test de l'existance de la valeur

            Try
                'récupère la valeur du premier élément
                valeur = doc.SelectNodes(xpath).ItemOf(index - 1).FirstChild.InnerText()
            Catch exx As System.NullReferenceException
                valeur = ""
            End Try

            'libère les ressources
            doc = Nothing

        Catch 'e As Exception
            'en cas d'erreur
            'Dim connectmsg As New Revo.connect
            'connectmsg.RevoLog(connectmsg.DateLog & "Cmd Get XML" & vbTab & "False" & vbTab & "Type: " & e.Message)
            'libère les ressources
            'doc = Nothing
            valeur = "" ' e.Message
        Finally
            '  valeur = ""
        End Try

        Return valeur

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ecrit la valeur pour le n-ième élément pointé par la requête path
    ''' Attention : Le premier noeud possède l'index 1
    ''' </summary>
    ''' <param name="xpath">Une requête XPath</param>
    ''' <param name="valeur">La valeur à écrire</param>
    ''' <param name="index">L'index de l'élément dans lequel écrire</param>
    ''' <remarks>
    ''' Utilisation |
    ''' exemple 1 : setElementValue("/polygone/point", "30;40", 5) spécifie que le 5ème élément "point" de "polygone" aura comme valeur "(30;40)" |
    ''' exemple 2 : setElementValue("/polygone/point[5]", "30;40") fait exactement la même chose.
    ''' </remarks>
    ''' <history>
    ''' 	[Gabriel]	10/05/2005	Created
    '''     [Gabriel]   06/06/2009  Updated
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub setElementValue(ByVal xpath As String, ByVal valeur As String, Optional ByVal index As Integer = 1)

        Try

            'charge le fichier xml
            Dim doc As New XmlDocument
            doc.Load(fichier)

            'definit la racine de l'element a modifier
            Dim parent As XmlNode = doc.SelectNodes(xpath).ItemOf(index - 1).ParentNode

            'définit le noeud
            Dim node As XmlNode = doc.SelectNodes(xpath).ItemOf(index - 1)

            'Si le noeud a des enfants alors il faut sauver les enfants
            If node.HasChildNodes Then

                'creation d'un nouvel element
                Dim clone As XmlElement = doc.CreateElement(node.Name)

                'on lui assigne la nouvelle valeur
                clone.InnerText = valeur

                'copie des enfants du noeud dans le nouvel element
                Dim child As XmlNode
                For Each child In node.ChildNodes
                    If child.GetType.ToString <> "System.Xml.XmlText" Then
                        clone.AppendChild(child.Clone())
                    End If
                    parent.ReplaceChild(clone, node)
                    node = clone
                Next
            Else  'sinon on change juste la valeur

                'creation d'un nouvel element
                Dim elem As XmlElement = doc.CreateElement(node.Name)

                'on lui assigne la nouvelle valeur
                elem.InnerText = valeur

                'on remplace par le nouveau noeud
                parent.ReplaceChild(elem, node)
            End If

            'Sauve les modifications
            doc.Save(fichier)

            doc = Nothing

        Catch e As Exception

            MsgBox("setSettings <> Erreur dans la modification de " & fichier & " : " & e.Message)

        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ajoute un element pointé par le n-ième noeud de la requête XPath
    ''' Attention : Le premier noeud possède l'index 1
    ''' </summary>
    ''' <param name="xpath">Une requête XPath</param>
    ''' <param name="nom">Le nom de l'élément à ajouter</param>
    ''' <param name="valeur">La valeur de l'élément à ajouter</param>
    ''' <param name="index">L'index de l'élément parent dans lequel on ajoute l'élément enfant "nom"</param>
    ''' <remarks>
    ''' Utilisation |
    ''' exemple 1 : addElement("/cd/piste", "titre", "54 Cymru beats", 5) ajoute un élément enfant "titre" au 5ème élément "piste", sa valeur sera "54 Cymru beats" |
    ''' exemple 2 : addElement("/cd/piste[5]", "titre", "54 Cymru beats") fait exactement la même chose.
    ''' </remarks>
    ''' <history>
    ''' 	[Gabriel]	10/05/2005	Created
    '''     [Gabriel]   06/06/2009  Updated
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub addElement(ByVal xpath As String, ByVal nom As String, ByVal valeur As String, Optional ByVal index As Integer = 1)

        Try
            'charge le fichier xml
            Dim doc As New XmlDocument
            doc.Load(fichier)


            'selectionne le noeud parent de l'élément à ajouter
            Dim root As XmlNode = doc.SelectNodes(xpath).ItemOf(index) '.ItemOf(index)
            Dim rootElement As XmlElement
            ' rootElement = doc.CreateElement("revo") '  root.ParentNod
            rootElement = doc.SelectSingleNode(xpath) '  root.ParentNode

            'création du nouvel élément
            Dim elem As XmlElement
            elem = doc.CreateElement(nom)
            elem.InnerText = valeur 'on lui assigne une valeur

            'puis on l'ajoute au noeud parent
            rootElement.AppendChild(elem)
            ' doc.DocumentElement.AppendChild(rootElement)

            'on sauvegarde
            doc.Save(fichier)

        Catch e As Exception
            'MsgBox("Erreur dans la création de l'élément : " & e.Message)

        End Try


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' supprime le n-ième element pointé par la requête XPath
    ''' Attention : Le premier noeud possède l'index 1
    ''' </summary>
    ''' <param name="xpath">Une requête XPath</param>
    ''' <param name="nom">Le nom de l'élément à supprimer</param>
    ''' <param name="index">L'index de l'élément à supprimer</param>
    ''' <remarks>
    ''' Utilisation |
    ''' exemple : deleteElement("/cd", "piste", 3) supprime le 3ème élément "piste"
    ''' </remarks>
    ''' <history>
    ''' 	[Gabriel]	10/05/2005	Created
    '''     [Gabriel]   06/06/2009  Updated
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub deleteElement(ByVal xpath As String, ByVal nom As String, Optional ByVal index As Integer = 1)

        Try
            'charge le fichier xml
            Dim doc As New XmlDocument
            doc.Load(fichier)

            'selectionne le noeud parent de l'élément à supprimer
            Dim root As XmlNode = doc.SelectSingleNode(xpath)

            'selectionne l'élément à supprimer
            Dim elem As XmlElement = doc.SelectNodes(xpath & "/" & nom).ItemOf(index - 1)

            'supprime l'élément
            root.RemoveChild(elem)

            'sauvegarde
            doc.Save(fichier)

        Catch e As Exception
            'MsgBox("Erreur dans la suppression de l'élément : " & e.Message)
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' supprime le 1er élément pointé par la requête XPath
    ''' Attention : Le premier noeud possède l'index 1
    ''' </summary>
    ''' <param name="xpath">Une requête XPath</param>
    ''' <remarks>
    ''' Utilisation |
    ''' exemple 1 : deleteElement("/cd/piste[3]) supprime le 3ème élément "piste"
    ''' exemple 2 : deleteElement("/cd/piste) supprime la première piste uniquement
    ''' </remarks>
    ''' <history>
    ''' 	[Gabriel]	10/05/2005	Created
    '''     [Gabriel]   06/06/2009  Updated
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub deleteElement(ByVal xpath As String)

        Try
            'charge le fichier xml
            Dim doc As New XmlDocument
            doc.Load(fichier)

            'selectionne le noeud parent de l'élément à supprimer
            Dim root As XmlNode = doc.SelectSingleNode(xpath).ParentNode

            'selectionne l'élément à supprimer
            Dim elem As XmlElement = doc.SelectSingleNode(xpath)

            'supprime l'élément
            root.RemoveChild(elem)

            'sauvegarde
            doc.Save(fichier)

        Catch e As Exception
            MsgBox("Erreur dans la suppression de l'élément : " & e.Message)
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ajoute un attribut "nom" de valeur "valeur" n-ième élément pointé par la requête XPath
    ''' Attention : Le premier noeud possède l'index 1
    ''' </summary>
    ''' <param name="xpath">Une requête XPath</param>
    ''' <param name="nom">Le nom de l'attribut à ajouter</param>
    ''' <param name="valeur">La valeur de l'attribut à ajouter</param>
    ''' <param name="index">L'index de l'élément auquel on ajoute l'attribut</param>
    ''' <remarks>
    ''' Utilisation |
    ''' exemple 1 : addAttribute("/cd/piste","numero", "3", 5) ajoute l'attribut "numero" de valeur "3" au 5ème élément "piste" du "cd"
    ''' exemple 2 : addAttribute("/cd/piste[5]","numero", "3") fait la même chose.
    ''' </remarks>
    ''' <history>
    ''' 	[Gabriel]	10/05/2005	Created
    '''     [Gabriel]   06/06/2009  Updated
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub addAttribute(ByVal xpath As String, ByVal nom As String, ByVal valeur As String, Optional ByVal index As Integer = 1)

        Try
            'charge le fichier xml
            Dim doc As New XmlDocument
            doc.Load(fichier)

            'declaration de l'élément auquel on ajoutera un attribut et de l'attribut en question
            Dim root As XmlNodeList = doc.SelectNodes(xpath)
            Dim attrib As XmlAttribute = doc.CreateAttribute(nom)

            'on valorise l'attribut
            attrib.InnerText = valeur

            'on ajoute l'attribut à l'élément
            root.ItemOf(index - 1).Attributes.Append(attrib)

            'et on sauvegarde
            doc.Save(fichier)

        Catch e As Exception
            MsgBox("Erreur dans la création de l'attribut : " & e.Message)
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Affecte à l'attribut "nom" la valeur "valeur" au n-ième élément pointé par la requête "path"
    ''' Attention : Le premier noeud possède l'index 1
    ''' </summary>
    ''' <param name="xpath">Une requête XPath</param>
    ''' <param name="nom">Le nom de l'attribut à affecter</param>
    ''' <param name="valeur">La nouvelle valeur de l'attribut</param>
    ''' <param name="index">L'index de l'élément auquel on affecte la nouvelle valeur de l'attribut</param>
    ''' <remarks>
    ''' Utilisation |
    ''' exemple 1 : setAttribute("/cd/piste","numero", "3", 5) définit l'attribut "numero" du 5ème élément "piste" de "cd" à la valeur "3"
    ''' exemple 2 : setAttribute("/cd/piste[5]","numero", "3") fait exactement la même chose
    ''' </remarks>
    ''' <history>
    ''' 	[Gabriel]	10/05/2005	Created
    '''     [Gabriel]   06/06/2009  Updated
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub setAttribute(ByVal xpath As String, ByVal nom As String, ByVal valeur As String, Optional ByVal index As Integer = 1)

        Try
            'on charge le fichier xml
            Dim doc As New XmlDocument
            doc.Load(fichier)

            'déclarations
            Dim root As XmlNode = doc.SelectNodes(xpath).ItemOf(index - 1)
            Dim attrib As XmlAttribute = doc.CreateAttribute(nom)

            'valorisation de l'attribut
            attrib.InnerText = valeur

            'ajout de l'attribut à l'élément
            root.Attributes.Append(attrib)

            'sauvegarde
            doc.Save(fichier)

        Catch e As Exception
            MsgBox("Erreur dans la modification de l'attribut : " & e.Message)
        End Try

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Retourne la valeur de l'attribut "nom" du n-ième élément pointé par la requête "XPath"
    ''' Attention : Le premier noeud possède l'index 1
    ''' </summary>
    ''' <param name="xpath">Une requête XPath</param>
    ''' <param name="nom">Le nom de l'attribut dont la valeur sera retournée</param>
    ''' <param name="index">L'index de l'élément auquel appartient l'attribut</param>
    ''' <returns>La valeur de l'attribut sélectionné</returns>
    ''' <remarks>
    ''' Utilisation |
    ''' exemple 1 : getAttribute("/cd/piste","numero", 5) retourne la valeur de l'attribut "numero" du 5ème élément "piste" de "cd"
    ''' exemple 2 : getAttribute("/cd/piste[5]","numero") fait la même chose.
    ''' </remarks>
    ''' <history>
    ''' 	[Gabriel]	10/05/2005	Created
    '''     [Gabriel]   06/06/2009  Updated
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function getAttribute(ByVal xpath As String, ByVal nom As String, Optional ByVal index As Integer = 1) As String

        Dim valeur As String = String.Empty
        Try
            'charge le fichier xml
            Dim doc As New XmlDocument
            doc.Load(fichier)

            'declarations
            Dim root As XmlNode = doc.SelectNodes(xpath).ItemOf(index - 1)
            Dim attrib As XmlAttribute = root.Attributes.GetNamedItem(nom)

            'récupération de la valeur de l'attribut
            valeur = attrib.InnerText()

        Catch e As Exception
            MsgBox("Erreur dans le retour de l'attribut : " & e.Message)
        End Try

        'retourne la valeur de l'attribut
        Return valeur

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ajoute l'element racine "nom"
    ''' </summary>
    ''' <param name="nom">Le nom de l'élément racine qui sera créé</param>
    ''' <remarks>
    ''' Utilisation | exemple : createRoot("Bibliothèque")
    ''' </remarks>
    ''' <history>
    ''' 	[Gabriel]	10/05/2005	Created
    '''     [Gabriel]   06/06/2009  Updated
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub createRoot(ByVal nom As String, Optional ByVal encoding As String = "ISO-8859-1")

        Try
            'déclare un nouveau document xml
            Dim doc As New XmlDocument

            'lui ajoute son entête et la balise racine
            doc.LoadXml("<?xml version='1.0' encoding='" & encoding & "'?>" & _
                         "<" & nom & ">" & _
                         "</" & nom & ">")

            'sauvegarde les modifications
            doc.Save(fichier)

        Catch e As Exception
            MsgBox("Erreur dans la création de la racine : " & e.Message)
        End Try
    End Sub
    '''
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Crée un nouveau fichier vierge
    ''' </summary>
    ''' <param name="nomFichier">[Chemin et ] Nom du fichier</param>
    ''' <remarks>
    ''' Veiller à ce que le fichier n'existe pas déjà sous peine d'écrasement de l'existant
    ''' </remarks>
    ''' <history>
    ''' 	[Gabriel]	10/05/2005	Created
    '''     [Gabriel]   06/06/2009  Updated
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub createNewFile(ByVal nomFichier As String)

        Try
            'creation d'un nouveau fichier
            System.IO.File.Create(nomFichier)

            System.Windows.Forms.Application.DoEvents()

        Catch e As Exception
            MsgBox("Erreur dans la création du fichier : " & e.Message)
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Permet de renvoyer sous forme de chaine le contenu du fichier XML.
    ''' </summary>
    ''' <returns>Une chaine de caractères</returns>
    ''' <remarks>
    ''' Le document contient des retours à la ligne pour une meilleure lecture
    ''' </remarks>
    ''' <history>
    ''' 	[Gabriel]	10/05/2005	Created
    '''     [Gabriel]   06/06/2009  Updated
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function getFormatedXMLString() As String

        'charge le fichier xml
        Dim doc As New XmlDocument
        doc.Load(fichier)

        'retourne le contenu intégral du fichier dans une chaine
        Return (doc.OuterXml).Replace("><", ">" & vbNewLine & "<")

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renvoie l'index de l'élément pointé par XPath et de valeur valeur, -1 si pas trouvé
    ''' Attention : Le premier noeud possède l'index 1
    ''' </summary>
    ''' <param name="xpath">Une requête XPath</param>
    ''' <param name="valeur">La valeur à rechercher</param>
    ''' <returns>Un entier</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Gabriel]	10/05/2005	Created
    '''     [Gabriel]   06/06/2009  Updated
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function getIndexOfElementContaining(ByVal xpath As String, ByVal valeur As String) As Integer
        Dim num As Integer = -1
        Try
            'charge le fichier xml
            Dim doc As New XmlDocument
            doc.Load(fichier)

            'déclarations
            Dim nodes As XmlNodeList = doc.SelectNodes(xpath)
            Dim nod As XmlElement
            Dim trouve As Boolean = False

            'recherche de l'élément
            For Each nod In nodes
                num = num + 1
                'si la valeur de l'élément correspond à la valeur cherchée alors
                If nod.InnerText = valeur Then
                    trouve = True
                    num += 1
                    Exit For
                End If
            Next

            'si non trouvé num <-- -1
            If Not trouve Then num = -1

            'on retourne la position
        Catch e As Exception
            MsgBox("Erreur de retour de l'index : " & e.Message)
        End Try
        Return num
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renvoie le nombre d'éléments pointés par XPath
    ''' </summary>
    ''' <param name="xpath"></param>
    ''' <returns>Un entier</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Gabriel]	10/05/2005	Created
    '''     [Gabriel]   06/06/2009  Updated
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function countElements(ByVal xpath As String) As Integer

        Dim nb As Integer = 0
        Try
            'Charge le fichier xml
            Dim doc As New XmlDocument
            doc.Load(fichier)

            'Compte les éléments correspondants au chemin
            nb = doc.SelectNodes(xpath).Count


        Catch e As Exception
            MsgBox("Erreur dans le comptage des éléments : " & e.Message)
            nb = -1
        End Try

        'retourne la valeur
        Return nb

    End Function

End Class

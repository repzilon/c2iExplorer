Attribute VB_Name = "Graphics"
'Salut !
'
'Ce module contient quelques fonctions graphiques pour Visual basic.
'Elles ont été écrites en VB 5 mais devraient fonctionner à partir de la
'version 4. Pour la version 3, le remplissage de feuille avec une image
'ne fonctionnera sûrement pas car il utilise une variable de type Picture,
'ce type n'existant pas (d'après Microsoft) dans la version 3.
'Voici le contenu de ce module :
'       1. Collision
'           Cette fonction permet de déterminer si deux objets se superposent.
'           Dans ce cas, elle renvoie True, sinon elle renvoie False.
'           Elle requiert 2 arguments, les 2 objets à tester. Il faut qu'ils soient
'           chargés, ou du moins la feuille qui les contient. L'ordre ne change rien.
'           NOTE : sauf coincidence, une feuille et un contrôle sur cette feuille
'           ne sont pas en collision, les coordonnées de la première étant
'           relatives à l'écran, et l'autre à la feuille.
'       2. FillForm
'           Cette fonction remplit une feuille avec une image, qui sera mise en
'           mosaïque si nécessaire. On lui donne deux arguments : la feuille et
'           l'image. L'image doit être chargée, et non un nom de fichier. Par
'           exemple : FillForm Form1, LoadPicture ("IMAGE.BMP"). Vous pouvez
'           aussi mettre pour le deuxième argument la propriété Picture d'un
'           quelconque objet.
'       3. FillPicture
'           Cette fonction fait exactement la même chose, mais avec un objet
'           PictureBox.
'       4. Dégradé
'           Cette fonction crée un dégradé dans l'objet spécifié. Il faut lui spécifier
'           l'objet, la couleur et la direction.
'           - Couleurs : il faut spécifier les trois composantes, à savoir le Rouge,
'           Vert et le Bleu. Chaque argument-couleur peut prendre True ou False.
'           Si, par exemple, on met du Vert et du Rouge, on aura un dégradé Jaune.
'           Le dernier argument spécifie le sens de variation. False pour avoir un
'           dégradé de la couleur au noir, True pour l'inverse.
'           - Direction : deux arguments sont concernés. Horizontal et vertical. Ils
'           sont eux aussi de type Boolean. S'ils sont tous les deux vrais, on aura
'           un dégradé diagonal.
'           Syntaxe : Dégradé (Rouge, Vert, Bleu, Horizontal, Vertical, Sens)
'           Valeurs par défaut :
'               - Rouge = False
'               - Vert = False
'               - Bleu = True
'               - Horizontal = True
'               - Vertical = False
'               - Sens = False
'
'Ouf ! Bon, je vous souhaite de bonnes créations grâce à ça... Encore deux trucs :
'       - Mettre de préférence les fonctions Fill[OBJET] dans l'évennement Paint
'       de l'objet concerné.
'       - NE PAS en faire autant pour Dégradé : en effet, en traçant ce dégradé,
'       l'évennement Paint est appelé ; il y aurait donc boucle infinie ! Il vaut mieux
'       le mettre dans le Resize d'une feuille...
'
'Tout ce que je vous demande, c'est de m'envoyer votre avis sur la chose. Vous
'pouvez le modifier à votre gré, mais faites-moi part de vos modifications, pour
'que je les mette dans les versions suivantes !
'
'A bientôt
'
'Nicolas JOYARD <ovalprom@club-internet.fr>

'je ne sais pas si c'est utile, mais on ne sait jamais...
#If Win32 Then
    DefLng A-Z
#Else
    DefInt A-Z
#End If

Sub Dégradé(obj As Object, Optional vRed As Variant, Optional vGreen As Variant, Optional vBlue As Variant, Optional vVert As Variant, Optional vHoriz As Variant, Optional vLightToDark As Variant)
    ' Donne des valeures par défaut aux paramètres optionnels
    If IsMissing(vRed) Then vRed = False
    If IsMissing(vBlue) Then vBlue = False
    If IsMissing(vGreen) Then vGreen = False
    If Not vRed And Not vGreen Then vBlue = True
    If IsMissing(vVert) Then vVert = False
    If IsMissing(vHoriz) Then vHoriz = Not vVert
    If Not vVert And Not vHoriz Then vHoriz = True
    If IsMissing(vLightToDark) Then vLightToDark = True


    On Error Resume Next
    With obj

        Dim fAutoRedraw As Boolean, ordDrawStyle As Integer
        Dim ordDrawMode As Integer, iDrawWidth As Integer
        Dim ordScaleMode As Integer
        Dim rScaleWidth As Single, rScaleHeight As Single
        fAutoRedraw = .AutoRedraw: iDrawWidth = .DrawWidth
        ordDrawStyle = .DrawStyle: ordDrawMode = .DrawMode
        rScaleWidth = .ScaleWidth: rScaleHeight = .ScaleHeight
        ordScaleMode = .ScaleMode
        If Err Then Exit Sub
        On Error GoTo 0
        fAutoRedraw = .AutoRedraw

        .AutoRedraw = True
        .DrawWidth = 2
        .DrawStyle = vbInsideSolid: .DrawMode = vbCopyPen
        .ScaleMode = vbPixels
        .ScaleWidth = 256 * 2: .ScaleHeight = 256 * 2

        Dim clr As Long, i As Integer, X As Integer, Y As Integer
        Dim iRed As Integer, iGreen As Integer, iBlue As Integer
        For i = 0 To 255
            If vLightToDark Then
                If vRed Then iRed = 255 - i
                If vBlue Then iBlue = 255 - i
                If vGreen Then iGreen = 255 - i
            Else
                If vRed Then iRed = i
                If vBlue Then iBlue = i
                If vGreen Then iGreen = i
            End If
            clr = RGB(iRed, iGreen, iBlue)
            If vVert Then
                obj.Line (0, Y)-(.ScaleWidth, Y + 2), clr, BF
                Y = Y + 2
            End If
            If vHoriz Then
                obj.Line (X, 0)-(X + 2, .ScaleHeight), clr, BF
                X = X + 2
            End If
        Next
        .AutoRedraw = fAutoRedraw: .DrawWidth = iDrawWidth
        .DrawStyle = ordDrawStyle: .DrawMode = ordDrawMode
        .ScaleMode = ordScaleMode
        .ScaleWidth = rScaleWidth: .ScaleHeight = rScaleHeight
    End With
End Sub
Public Function FillForm(Feuille As Form, Image As Picture) As Boolean

    'NE PRODUIT RIEN DU TOUT SI UN ARGUMENT EST INCORRECT OU
    'S'IL Y A UNE ERREUR !
    FillForm = True

    On Error GoTo Fin
    Dim intX As Integer
    Dim intY As Integer
    Dim sngWidth As Single
    Dim sngHeight As Single
    sngWidth = Image.Width / 2
    sngHeight = Image.Height / 2
    For intX = 0 To Int(Feuille.ScaleWidth / sngWidth)
        For intY = 0 To Int(Feuille.ScaleHeight / sngHeight)
            Feuille.PaintPicture Image, intX * sngWidth, intY * sngHeight, sngWidth, sngHeight, 0, 0
        Next
    Next
    Exit Function

Fin:
    FillForm = False
    MsgBox Erl
End Function


Sub FillPicture(PicBox As PictureBox, Image As Picture)

    'NE PRODUIT RIEN DU TOUT SI UN ARGUMENT EST INCORRECT OU
    'S'IL Y A UNE ERREUR !

    On Error GoTo 100
    Dim intX As Integer
    Dim intY As Integer
    Dim sngWidth As Single
    Dim sngHeight As Single
    sngWidth = Image.Width / 2
    sngHeight = Image.Height / 2
    For intX = 0 To Int(PicBox.ScaleWidth / sngWidth)
        For intY = 0 To Int(PicBox.ScaleHeight / sngHeight)
            PicBox.PaintPicture Image, intX * sngWidth, intY * sngHeight, sngWidth, sngHeight, 0, 0
        Next
    Next
100
End Sub


Function Collision(Objet1 As Object, Objet2 As Object) As Boolean
    'La fonction Collision envoie True si une partie de l'objet1
    'superpose une partie de l'objet2 ou inversement, sinon
    'renvoie False

    On Error GoTo errhandle
    Dim o1 As Object, o2 As Object
    Set o1 = Objet1
    Set o2 = Objet2

    'calcul des coordonnées de O1
    With o1
        'a = x1     b = y1 (coin haut-gauche)
        'c = x2     d = y2 (coin bas-droite)
        A = .Left
        b = .Top
        c = .Left + .Width
        d = .Top + .Height
    End With

    'calcul des coordonnées de O2
    With o2
        'e = x1     f = y1 (coin haut-gauche)
        'g = x2     h = y2 (coin bas-droite)
        e = .Left
        f = .Top
        g = .Left + .Width
        h = .Top + .Height
    End With

    'Vérification des 4 possibilités de superposition. Si aucune n'est
    'vérifiée, O1 et O2 ne sont pas en collision.


    If (e <= c And e >= A) And (f <= d And f >= b) Then
        Collision = True
        Exit Function
    End If
    If (e <= c And e >= A) And (h <= d And h >= b) Then
        Collision = True
        Exit Function
    End If
    If (g <= c And e >= g) And (h <= d And h >= b) Then
        Collision = True
        Exit Function
    End If
    If (g <= c And g >= A) And (f <= d And f >= b) Then
        Collision = True
        Exit Function
    End If
    If (e < A And b < f) And (c < g And h < d) Then
        Collision = True
        Exit Function
    End If


    'on inverse les 2 objets pour calculer les cas inverses
    Set o2 = Objet1
    Set o1 = Objet2

    'calcul des coordonnées de O1
    With o1
        'a = x1     b = y1 (coin haut-gauche)
        'c = x2     d = y2 (coin bas-droite)
        A = .Left
        b = .Top
        c = .Left + .Width
        d = .Top + .Height
    End With

    'calcul des coordonnées de O2
    With o2
        'e = x1     f = y1 (coin haut-gauche)
        'g = x2     h = y2 (coin bas-droite)
        e = .Left
        f = .Top
        g = .Left + .Width
        h = .Top + .Height
    End With

    'Vérification des 4 possibilités de superposition. Si aucune n'est
    'vérifiée, O1 et O2 ne sont pas en collision.


    If (e <= c And e >= A) And (f <= d And f >= b) Then
        Collision = True
        Exit Function
    End If
    If (e <= c And e >= A) And (h <= d And h >= b) Then
        Collision = True
        Exit Function
    End If
    If (g <= c And e >= g) And (h <= d And h >= b) Then
        Collision = True
        Exit Function
    End If
    If (g <= c And g >= A) And (f <= d And f >= b) Then
        Collision = True
        Exit Function
    End If
    If (e < A And b < f) And (c < g And h < d) Then
        Collision = True
        Exit Function
    End If


    'Si ce point est exécuté, il n'y a pas collision.
    Collision = False
    Exit Function

errhandle:
    Collision = False
    'erreur. Objet(s) non chargé(s).
End Function




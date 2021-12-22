' Ecrire un script qui calcule la surface d'un cercle
' L'utilisateur doit donner la valeur du rayon

Const M_PI = 3.14

Function surface_cercle(R)
      surface_cercle = M_PI * R * R
End Function

WScript.StdOut.Write("Donne moi le rayon du cercle : ")
rayon = WScript.StdIn.ReadLine()

aire = surface_cercle(rayon)
WScript.StdOut.WriteLine("L'aire du cercle est : " & aire)



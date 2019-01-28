wget -r "http://127.0.0.1/Gallery/Color_Ranges/Colors.asp?&Type.html"
wget -r -np "http://127.0.0.1/CDmenu.asp"
wget -r -np "http://127.0.0.1/Gallery/GenerateCD.asp"



wget -r -np "http://127.0.0.1/Gallery/GenImages.asp"



Ren "C:\Documents and Settings\wally\My Documents\Generate CD Gallery\127.0.0.1\CDmenu.asp" "gallery.html"

XCopy "Files\*.*" "127.0.0.1\*.*" /H /E

echo "Finshed CD Files Generations"

pause
wget -r "http://127.0.0.1/Dev/Gallery/Color_Ranges/Colors.asp?&Type.html"
wget -r -np "http://127.0.0.1/Dev/CDmenu.asp"
wget -r -np "http://127.0.0.1/Dev/Gallery/GenerateCD.asp"



wget -r -np "http://127.0.0.1/Dev/Gallery/GenImages.asp"



Ren "D:\MyWorks\Generate CD Gallery\127.0.0.1\Dev\CDmenu.asp" "index.html"

XCopy "Files\*.*" "127.0.0.1\*.*" /H /E

echo "Finshed CD Files Generations"

pause
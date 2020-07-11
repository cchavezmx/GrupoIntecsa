option compare database


'guarda el valor de estado de ventana 

dim dwreturn as long 

'constante de estado de ventan 

const sw_hide = 0
const sw_shownormal  = 1
const sw_minimizada = 2
cosnt sw_maximizada = 3

'Se identifica el tipo de plataforma 

#if win64 then
    private declare ptrsafe fuction iswindowVisible lib "user32" (byVal hwnd as long)  as long 
    private declare ptrsafe fuction showwindow lib "user32" (byVal hwnd as long ncmdshow as long) as long     

#Else 
    private declare ptrsafe fuction iswindowVisible lib "user32" (byVal hwnd as long) as long 
    private declare ptrsafe fuction showwindow lib "user32" (byVal hwnd as long ncmdshow as long) as long
#end If 

'Llamada de funcio para ocultar ventan

if procedure = "hide" then 
    dwretur  = showwindow(application.hwndaccessApp, sw_hide)
end if
if procedure = "show" then
    dwretur  = showwindow(application.hwndaccessApp, sw_showmaximized)
End If
if procedure = "Minimized" then
    dwretur  = showwindow(application.hwndaccessApp, sw_showmaximized)
end if 

if switchstatus = true then 
  if iswindowsvisible(hwndaccessapp) = 1 then
  dwretur = showwindow(application.hwndaccessApp, sw_hide)
  else dwretur = showwindow(application.hwndaccessApp, sw_showmaximized)
  end If   
end if

if statuscheck = true then
    if iswindowsvisible(hwndaccessapp) = 0 then
        faccesswindows = false
    end If
    if iswindowsvisible(hwndaccessapp) = 0 then
       faccesswindows = false          
    End If
end If 

end Function 

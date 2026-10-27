Unicode true
!include "MUI2.nsh"

; 版本号和输出文件名由 生成安装包.ps1 从 版本.json 读取后传入，禁止在此硬编码。
!ifndef APPVERSION
  !error "缺少 APPVERSION。请运行 生成安装包.ps1，不要直接编译 安装程序.nsi。"
!endif
!ifndef OUTFILE
  !error "缺少 OUTFILE。请运行 生成安装包.ps1，不要直接编译 安装程序.nsi。"
!endif

Name "通道生效签名短链程序"
OutFile "${OUTFILE}"
InstallDir "$LOCALAPPDATA\Programs\通道生效签名短链程序"
RequestExecutionLevel user
SetCompressor /SOLID lzma
Icon "发布资源\应用图标.ico"
UninstallIcon "发布资源\应用图标.ico"
ShowInstDetails show
ShowUninstDetails show

VIProductVersion "${APPVERSION}.0"
VIAddVersionKey /LANG=2052 "ProductName" "通道生效签名短链程序"
VIAddVersionKey /LANG=2052 "ProductVersion" "${APPVERSION}"
VIAddVersionKey /LANG=2052 "FileVersion" "${APPVERSION}"
VIAddVersionKey /LANG=2052 "FileDescription" "通道生效签名短链程序安装包"
VIAddVersionKey /LANG=2052 "LegalCopyright" "本地使用"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "SimpChinese"

Section "安装程序" SecMain
  SetOutPath "$INSTDIR"
  File "构建临时目录\发布\通道生效签名短链程序.exe"
  File "构建临时目录\发布\AntdUI.dll"
  File /oname=AntdUI-LICENSE.txt "源码\第三方\AntdUI\LICENSE.txt"
  File "说明文档\安装版使用说明.md"
  File "说明文档\子端口分配使用说明.md"
  WriteUninstaller "$INSTDIR\卸载程序.exe"
  ReadEnvStr $0 "CHANNEL_TOOL_INSTALL_TEST"
  StrCmp $0 "1" install_done
  CreateDirectory "$SMPROGRAMS\通道生效签名短链程序"
  Delete "$SMPROGRAMS\通道生效签名短链程序\通道生效签名短链程序.lnk"
  Delete "$DESKTOP\通道生效签名短链程序.lnk"
  CreateShortCut "$SMPROGRAMS\通道生效签名短链程序\通道生效签名短链程序.lnk" "$INSTDIR\通道生效签名短链程序.exe" "" "$INSTDIR\通道生效签名短链程序.exe" 0
  CreateShortCut "$SMPROGRAMS\通道生效签名短链程序\卸载程序.lnk" "$INSTDIR\卸载程序.exe"
  CreateShortCut "$DESKTOP\通道生效签名短链程序.lnk" "$INSTDIR\通道生效签名短链程序.exe" "" "$INSTDIR\通道生效签名短链程序.exe" 0
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\ChannelSignatureTool" "DisplayName" "通道生效签名短链程序"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\ChannelSignatureTool" "DisplayVersion" "${APPVERSION}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\ChannelSignatureTool" "DisplayIcon" "$INSTDIR\通道生效签名短链程序.exe"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\ChannelSignatureTool" "UninstallString" "$\"$INSTDIR\卸载程序.exe$\""
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\ChannelSignatureTool" "InstallLocation" "$INSTDIR"
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\ChannelSignatureTool" "NoModify" 1
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\ChannelSignatureTool" "NoRepair" 1
  install_done:
SectionEnd

Section "Uninstall"
  Delete "$INSTDIR\通道生效签名短链程序.exe"
  Delete "$INSTDIR\AntdUI.dll"
  Delete "$INSTDIR\AntdUI-LICENSE.txt"
  Delete "$INSTDIR\安装版使用说明.md"
  Delete "$INSTDIR\子端口分配使用说明.md"
  Delete "$INSTDIR\卸载程序.exe"
  RMDir "$INSTDIR"
  Delete "$SMPROGRAMS\通道生效签名短链程序\通道生效签名短链程序.lnk"
  Delete "$SMPROGRAMS\通道生效签名短链程序\卸载程序.lnk"
  RMDir "$SMPROGRAMS\通道生效签名短链程序"
  Delete "$DESKTOP\通道生效签名短链程序.lnk"
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\ChannelSignatureTool"
SectionEnd

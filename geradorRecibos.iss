[Setup]
; Configurações básicas do instalador
AppName=Gerador de Recibos
AppVersion=1.0
DefaultDirName={autopf}\GeradorRecibos
DefaultGroupName=Gerador de Recibos
OutputDir=.
OutputBaseFilename=GeradorRecibosInstalador
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin

; Ícone do instalador
SetupIconFile=icon.ico

[Files]
; Adiciona o executável principal gerado pelo PyInstaller
Source: "dist\main\main.exe"; DestDir: "{app}"; Flags: ignoreversion

; Adiciona os recursos (ícones e outros assets)
Source: "dist\main\_internal\assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs

; Adiciona os templates (HTML e CSS)
Source: "dist\main\_internal\templates\*"; DestDir: "{app}\templates"; Flags: ignoreversion recursesubdirs createallsubdirs

; Adiciona outras dependências necessárias do PyInstaller
Source: "dist\main\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Cria um atalho no Menu Iniciar
Name: "{group}\Gerador de Recibos"; Filename: "{app}\main.exe"

; Cria um atalho na Área de Trabalho
Name: "{commondesktop}\Gerador de Recibos"; Filename: "{app}\main.exe"

[Messages]
; Mensagem ao concluir a instalação
FinishedLabel=Instalação concluída! O Gerador de Recibos está pronto para uso.

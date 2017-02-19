# launcheralankoba
Launcher com atualizador para Muonline.


launcher alankoba
--- 19/02/2017 ---
 * Abertura do código fonte

### Anotações do desenvolvedor
 Infelizmente o histórico de alterações original não existe mais. Descrevo abaixo todas as funcionalidades do launcher.

* Verificação CRC dos arquivos (algoritmo extremamente rápido)
* Compactação no formato RAR (utiliza a lib unrar.dll)
* Configurações do Muonline como usuário, ajustes de som e resolução...
* Compatível com Windows XP/Vista/7 (Testes foram feitos no Windows 8/10 e está tudo OK).
* Utiliza somente componentes nativos do Windows.
* Atualização do próprio launcher.
* Leve e livre de "false-positives".

###Como usar
* Abra o projeto no VB6.
* Configure os valores no form update.
* Menu File -> make jogar.exe.
* Na pasta do gerador de atualizações localize o arquivo config.ini e altere o valor da chave name.

###Dependências
jogar.exe
atualiza.exe
unrar.dll
 

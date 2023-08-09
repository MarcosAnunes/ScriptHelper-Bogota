# ScriptHelper-Bogota
Versão em espanhol do ScriptHelper

- LOGIN: Devem ser utilizadas as credenciais ADM do domínio (admUSER) 
* Não precisa do @dominio.net
* É possível usar o script sem logar, mas algumas funções não estarão disponíveis para uso. (Ativar e desativar conta e redefinir senha)
* Se o agente errar a senha ADM duas vezes, o script vai informar que chegou ao numero máximo de tentativa, e vai iniciar com as funções desabilitadas. 

- Pesquisa: Pode ser utilizado: Matricula, e-mail, login Heiway e nome do usuário (Pesquisa por nome esta em teste, não recomendo utilizar).
* Script valida se usuário é VIP imprime na tela "VIP" e muda a cor de exibição.
* Ao pesquisar as informações básicas do usuário já são copiadas para o clipboard. 

Funções:
- Desbloquear: Possui camada de proteção para evitar enganos, também é possível forçar o desbloqueio da conta do usuário. 
- Desativar e ativar a conta: Possui camada de proteção tanto para ATIVAR quanto para DESATIVAR para evitar enganos no uso. (Função só disponível se usuário fizer login ADM).
- Imprimir grupos: Possui função de busca na janela, sendo possível verificar por palavras chave.
- Redefinição de senha: (Função só disponível se usuário fizer login ADM).
- Buscador de gestor: Imprime as informações do gestor na tela, e copia para o clipboard.
- Alteração de telefone: Possível alterar os dois parâmetros chave de telefone no AD, mobilephone e telephone, este ultimo é o utilizado pelo BOT Boomi para encaminhar chamados para a fila dispatch.
- Copiar as informações: Copia as informações do usuário novamente para o clipboard.

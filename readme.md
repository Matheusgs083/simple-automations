### **Oque aprendi com esses códigos**:

## resumo

Meu primeiro contato com automação foi através do Excel, onde aprendi o básico de macros em VBA. Apesar de ser uma linguagem considerada datada, ela ainda é muito útil no ambiente corporativo. Isso foi especialmente verdade na empresa em que trabalhei, pois não tínhamos um sistema de gestão decente.

Por se tratar de um sistema legado, os relatórios gerados ou continham informações demais, ou vinham com dados de menos. Para otimizar meu tempo, precisei encontrar uma maneira de organizar esses relatórios. Mesmo conhecendo outras ferramentas mais simples e práticas, como o Power BI, optei pelo Excel por já saber da sua integração com Python.

Como eu não estava atuando na minha área, queria aproveitar a oportunidade para estudar algo aplicável. No fim, todo o processo de estudo foi divertido. Hoje, não me considero um programador VBA ou sequer um especialista em excel, mas consigo desenrolar.

## automações com python

A necessidade de automação surgiu com um novo problema na empresa: a tarefa diária e sequencial de alterar dados de "x" para "y". Como estudo programação, obviamente não faria isso manualmente.

Pesquisei por automações com Python e descobri várias maneiras, mas a mais adequada para o trabalho era o PyAutoGUI. É uma biblioteca com semântica e entendimento muito simples, que simula os movimentos do mouse e do teclado. No primeiro dia do problema, já criei uma planilha para organizar os dados — que também vinham de forma desorganizada — e, em seguida, desenvolvi o código.

O script era simples: ele lia os dados da planilha com as bibliotecas Pandas e OpenPyXL e os inseria no sistema legado. Apesar da simplicidade, seria quase impossível errar. O desafio era que o sistema tinha um tempo de resposta horrível e imprevisível, o que tornava difícil calcular o tempo exato entre um comando e outro.

Para contornar isso, a primeira solução foi tornar a execução mais lenta nas ações que exigiam carregamento de página. O próprio PyAutoGUI poderia ter resolvido isso de forma mais elegante, já que ele consegue procurar imagens na tela. Bastava tirar um print de um elemento e pedir para o script aguardar até que ele aparecesse. No entanto, o sistema não facilitava: a única coisa que atualizava na página eram os próprios valores que mudavam a cada inserção. Assim, a execução lenta acabou sendo a única solução viável.

## criação de mais automações e interfaces

As automações que desenvolvi ganharam notoriedade na empresa, o que me motivou a procurar maneiras de compartilhá-las com meus colegas. A primeira solução foi usar o Tkinter para criar interfaces gráficas, por ser uma biblioteca simples e já inclusa na instalação padrão do Python.

Embora as interfaces tenham sido criadas, um novo desafio surgiu: meus colegas ainda precisariam ter todas as outras bibliotecas instaladas em suas máquinas, o que não era uma tarefa simples para todos.

Foi então que descobri o PyInstaller. Com um único comando, ele me permitiu criar um arquivo executável que não exigia instalações complementares por parte do usuário. Apesar de o arquivo final ficar um pouco grande, essa se mostrou a melhor solução. Dessa forma, consegui distribuir as ferramentas e facilitar a rotina de muitos na equipe com automações simples de mouse e teclado.

## dificuldades

As minhas principais dificuldades foram o sistema de resposta lenta e a falta de tempo para implementar meus conhecimentos em programação. Tive também problemas com a instalação das bibliotecas e da IDE para rodar os códigos, já que tudo isso era novo na empresa e os computadores tinham proteções de segurança. Foi difícil convencê-los a me deixar prosseguir, mas a situação só mudou depois que viram o poder do PyAutoGUI."

## conclusão

Programação é muito divertido, tanto no meio acadêmico, como na vida real, mesmo em processo que eu não utilizava computadores, eu consegui aplicar métodos que eu aprendi na programação, por exemplo, algoritmos, onde eu sabia que se eu seguisse um passo a passo, eu faria muito mais rapido trabalhos "braçais", em resumo, tudo foi organização previa e planejamento. muito bom estudar programação.

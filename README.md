# Ferramenta Simples - Criador de Procurações, Declarações de Hipossuficiência e Contratos

## Ferramenta simples auxiliou nas minhas tarefas do dia a dia em meus trabalhos de assistente administrativo em um escritório de Advocacia /nA ideia originalmente veio a partir de uma série de problemas, como:
<ul>
	<li>Falta de agilidade na criação destes documentos</li>
	<li>Dificuldade na criação destes documentos na ausência de um computador</li>
<li>Recorrentes erros envolvendo Data, CPF e RG digitados errados, entre outras dificuldades</li>
</ul>

## Novidades

Agora é possível importar dados diretamente de uma imagem ou PDF do RG. Basta
clicar em **Carregar RG** na tela principal. A extração utiliza OpenCV e
Caso não queira alterar variáveis do sistema, crie uma pasta `tessdata` no
mesmo diretório deste projeto e coloque o arquivo `por.traineddata` nela. O
código ajustará automaticamente `TESSDATA_PREFIX` para esse caminho.

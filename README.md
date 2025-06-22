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
pytesseract para reconhecer os campos de nome, CPF e RG.

## Dependências

Instale os seguintes pacotes Python:

```
pip install opencv-python pytesseract pdf2image docx2pdf num2words
```

Além disso, é necessário ter o utilitário **Tesseract OCR** e o pacote
**poppler** (para converter PDFs em imagens) instalados no sistema.
Em alguns ambientes pode ser preciso configurar a variável de ambiente
`TESSDATA_PREFIX` apontando para o diretório `tessdata` do Tesseract.
Certifique-se também de que o arquivo `por.traineddata` (idioma português)
esteja presente nesse diretório.

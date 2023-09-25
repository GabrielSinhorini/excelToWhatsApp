<h1>Automatização para WhatsApp</h1>

<h3>O que é? Como funciona?</h3>

<p>Automatização para envio de mensagens via WhatsApp, chamando os números pela extensão <strong>WA Web Plus</strong>, pegando os dados de uma planilha Excel, possuindo uma coluna com o <strong>Telefone</strong> e outra coluna com o <strong>Texto</strong>. Possui uma interface gráfica simples para selecionar o método de envio, importar a planilha e ajustar a velocidade do envio das mensagens.</p>

<br>

<h3>Métodos de envio</h3>

<p>É possível selecionar dois tipos de envio: o método por <strong>Imagem</strong> e por <strong>Atalho</strong>.</p>
<p><strong>O Método por Imagem</strong> tem o funcionamento referente ao reconhecimento da imagem de adicionar contato, fornecido pela extensão adicionada. Após clicado, o número é colado para a abertura da conversa e o envio da mensagem, porém antes de passar para a etapa do envio da mensagem, é feito uma nova verificação se caso a imagem de erro é informada. Se caso o erro for apresentado, o programa saí da tela de erro para continuar a execução normalmente, pulando o envio da linha em questão. O método de imagem é o mais recomendado e fornece o relatório no console no final da execução.</p>
<p><strong>Método por Atalho</strong> faz a execução de todo o programa sem nenhum reconhecimento de imagem, apenas utilizando o atalho CTRL+ALT+S para adicionar o número e enviar a mensagem novamente. No final de toda execução, é solicitado o ESC duas vezes para prevenir o erro de números inválidos.</stonrg></p>
<br>

<h3>Bibliotecas utilizadas</h3>
<ul>
    <li>Pandas</li>
    <li>PyautoGUI</li>
    <li>TKinter</li>
    <li>ClipBoard</li>
</ul>

<br>

<p>Baixe a última versão <a href="https://github.com/GabrielSinhorini/excelToWhatsApp/releases/tag/v0.5-alpha">Aqui</a></p>

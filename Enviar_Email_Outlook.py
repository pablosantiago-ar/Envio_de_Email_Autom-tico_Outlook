#Para baixar a bliblioteca, rode o comando - pip install pywin32

#Importa a Biblioteca Win32, que dá acesso a diversas APIs do windows
import win32com.client as win32

#Cria a integração com o Outlook. OBS: O programa só funcionará com o Outlook versão desktop, logado na conta que enviará os emails.
outlook = win32.Dispatch('outlook.application')

#Cria um novo email
email = outlook.CreateItem(0)

#Entre as aspas, colocar o email dos destinatários, separando-os por ';'. Exemplo: "example@gmail.com;example2@gmail.com"
#Ao copiar dados de email do Excel, os mesmos viram com quebra de linha, o qual não funcionará no código.
#Para retirar a quebra de linha e adicionar o caractere separador ';', ir ao site https://www.4devs.com.br/remover_trocar_quebra_linha
#Após isso, colar os emails no espaço destinado, e escolher para separar por caractere ;. copie o resultado e cole dentro das aspas do email.to
#O limite de destinatários por mensagem é de 500 usuários!
email.to ="destinatário"

#Assunto do email
email.Subject = "teste automação"

#Texto de dentro do email, escrever a mensagem com base em HTML.
email.HTMLBody = """
<h1>O furo esmagador</h1>

<p>Por Chris Mills</p>

<h2>Capítulo 1: A noite escura</h2>

<p>
  Era uma noite escura. Em algum lugar, uma coruja piou. A chuva caiu no chão
  ...
</p>

<h2>Capítulo 2: O eterno silêncio</h2>

<p>
  Nosso protagonista não podia ver mais que um relance da figura sombria ...
</p>

<h3>O espectro fala</h3>

<p>
  Várias horas se passaram, quando, de repente, o espectro ficou em pé e
  exclamou: "Por favor, tenha piedade da minha alma!"
</p>
"""
#Declarar a uma variável anexo, atribuindo a ela o destino do arquivo que deseja anexar
anexo = "C://Users/Pablo Rodrigues/Documents/enviaremail.py"
email.Attachments.Add(anexo)

#Envia o email
email.Send()
print("Email enviado")
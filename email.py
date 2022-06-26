import win32com.client as win32


var = 10

def enviar_email():
    # criar a integração com o outlook
    outlook = win32.Dispatch('outlook.application')
    
    # criar um email
    email = outlook.CreateItem(0)
    # configurar as informações do seu e-mail
    email.To = "bn.eletrica@gmail.com"
    email.Subject = "E-mail automático do Python"
    email.HTMLBody = f"""
    <!--corpo HTML-->
    <body style="background-color:rgb(78, 38, 110); margin: auto; width:50%; justify-content:center; align-items:center;">
    <section>
        <div class="container">
            <h1>PROGAMA FEITO EM PYTHON<h1/>          
            </div>
        </div>
    </section>
    <footer style="text-align: center; margin-top: 670px; color: white;">
        <p>Created by - Helysson</p>
    </footer>
    </body>
    """

    # envio de arquivo, em anexo.
    #anexo = "c:/Users/helys/OneDrive/Área de Trabalho/test/index.html"
    # email.Attachments.Add(anexo)

    email.Send()  # função de envio de email.
    print("Email Enviado!")


enviar_email()

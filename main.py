#importando pacotes para envio de emails
import win32cpm.client as win32





# condição criada para envio de emails de juiz de fora
emailsJF= ['lista de emails']
count = 0

for i in range():#dentro do range, vai a quantidades de repetições seu for vai fazer
    #Envio email Juiz de Fora
    #integração com outlook
    outlook = win32.Dispatch('outlook.application')

    #criação do email
    emailJF = outlook.CreateItem(0)

    #configurar as informaçoes do email 
    emailJF.To = f"destino@email.com;{emailsJF[count]}"#inserindo email destino@email.com como email em copia e usando emailsJF[count] que passara pela lista de emails
    emailJF.Subject = "Disponibilidade de horário Novembro"
    #formatado com linguagem html
    emailJF.HTMLBody = """ 
        <h5>Aqui vai o corpo do email</h5><br>
        
        <h1>Pode ser configurado de acordo com HTML</h1>
        
        <h2>Para uso do codigo estar logado no outlook, pois a integração de envios e feito de la</h2>



    """
    #envio do email
    emailJF.Send()
    #adicionando valor ao contador
    count = count +1
    

################################################################ FIM ENVIO EMAILS JF #################################################################


#####################################################Envio email Uberlandia#########################################################


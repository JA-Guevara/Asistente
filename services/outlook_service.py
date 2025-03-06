# services/outlook_service.py
import os
from decouple import config
from exchangelib import Credentials, Account, Configuration, DELEGATE
from exchangelib.errors import TransportError, AutoDiscoverFailed

def connect_to_email():
    try:
        # Cargar variables de entorno
        email_address = config('EMAIL_ADDRESS')
        email_password = config('EMAIL_PASSWORD')
        email_server = 'outlook.office365.com'  # Servidor de Office 365
        ews_url = 'https://outlook.office365.com/EWS/Exchange.asmx'  # Endpoint de EWS

        # Configurar credenciales
        credentials = Credentials(email_address, email_password)
        email_config = Configuration(service_endpoint=ews_url, credentials=credentials)

        # Conectar a la cuenta
        account = Account(primary_smtp_address=email_address, config=email_config,
                          autodiscover=False, access_type=DELEGATE)
        return account
    except TransportError as e:
        print(f"Error de transporte: {e}")
        return None
    except AutoDiscoverFailed as e:
        print(f"Error en la autodetección: {e}")
        return None
    except Exception as e:
        print(f"Error inesperado: {e}")
        return None

def get_emails(account, num_emails=10):
    if account is None:
        print("No se pudo conectar a la cuenta de correo.")
        return

    # Obtener los últimos correos
    for item in account.inbox.all().order_by('-datetime_received')[:num_emails]:
        print(f"De: {item.sender.email_address}")
        print(f"Asunto: {item.subject}")
        print(f"Fecha: {item.datetime_received}")
        print(f"Cuerpo: {item.body[:100]}...")  # Muestra los primeros 100 caracteres del cuerpo
        print("-" * 40)

if __name__ == "__main__":
    account = connect_to_email()
    if account:
        get_emails(account)
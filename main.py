# main.py
from services.outlook_service import connect_to_email, get_emails

def main():
    print("Conectando a la bandeja de correo...")
    account = connect_to_email()
    print("Conexión exitosa. Obteniendo los últimos correos...")
    get_emails(account, num_emails=5)  # Obtener los últimos 5 correos

if __name__ == "__main__":
    main()
from tkinter import messagebox
import customtkinter as ctk
from tkinter import filedialog
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

class App:
    def __init__(self, root):
        self.root = root
        self.archivo_excel = None
        self.archivo_a_enviar = None
        self.data_frame = None
        self.root.title("Autenticación")
        self.root.geometry("400x300")

        ctk.CTkLabel(root, text="Email:").pack(pady=5)
        self.entry_email = ctk.CTkEntry(root)
        self.entry_email.pack(pady=5)

        ctk.CTkLabel(root, text="Contraseña:").pack(pady=5)
        self.entry_contraseña = ctk.CTkEntry(root, show="*")
        self.entry_contraseña.pack(pady=5)

        ctk.CTkButton(root, text="Iniciar sesión", command=self.verificar_credenciales).pack(pady=20)

    def verificar_credenciales(self):
        email = self.entry_email.get()
        contraseña = self.entry_contraseña.get()

        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email, contraseña)
            server.quit()
            self.abrir_ventana_envio(email, contraseña)
        except smtplib.SMTPAuthenticationError:
            ctk.CTkLabel(self.root, text="Autenticación fallida. Inténtalo de nuevo.", fg_color="red").pack(pady=5)

    def abrir_ventana_envio(self, email, contraseña):
        self.root.destroy()

        # ---- CREACION VENTANA NUEVA ------

        self.nueva_ventana = ctk.CTk()
        self.nueva_ventana.title("Enviar correo")
        self.nueva_ventana.geometry("1000x700")

        # ---------------

        ctk.CTkLabel(self.nueva_ventana, text="Desde (Mail desde donde se enviará):").pack(pady=5)
        self.entry_from = ctk.CTkEntry(self.nueva_ventana, width=350)
        self.entry_from.pack(pady=3)

        ctk.CTkLabel(self.nueva_ventana, text="Asunto:").pack(pady=5)
        self.entry_asunto = ctk.CTkEntry(self.nueva_ventana, width=700)
        self.entry_asunto.pack(pady=3)

        ctk.CTkLabel(self.nueva_ventana, text="Cuerpo del mail:").pack(pady=5)
        self.textbox_cuerpo = ctk.CTkTextbox(self.nueva_ventana, width=700, height=250)
        self.textbox_cuerpo.pack(pady=3)

        ctk.CTkButton(self.nueva_ventana, width=400, text="Adjuntar Excel De Destinatarios", command=self.adjuntar_archivo).pack(pady=30)

        ctk.CTkButton(self.nueva_ventana, width=400, text="Adjuntar Archivo Común", command=self.adjuntar_archivo_comun).pack(pady=10)  

        ctk.CTkButton(self.nueva_ventana, width=400, text="ENVIAR", command=lambda: self.enviar_correo(email, contraseña)).pack(pady=30)

        self.nueva_ventana.mainloop()

    def adjuntar_archivo(self):
        self.archivo_excel = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.archivo_excel:
            self.data_frame = pd.read_excel(self.archivo_excel)
            ctk.CTkLabel(self.nueva_ventana, text="Archivo Excel cargado", fg_color="green").pack(pady=5)
        else:
            self.data_frame = None
            ctk.CTkLabel(self.nueva_ventana, text="No se seleccionó ningún archivo", fg_color="red").pack(pady=5)
    
    def adjuntar_archivo_comun(self):
        self.archivo_a_enviar = filedialog.askopenfilename(filetypes=[("Todos los archivos", "*.*")])
        if self.archivo_a_enviar:
            ctk.CTkLabel(self.nueva_ventana, text="Archivo común cargado", fg_color="green").pack(pady=5)
        else:
            self.archivo_a_enviar = None
            ctk.CTkLabel(self.nueva_ventana, text="No se seleccionó ningún archivo común", fg_color="red").pack(pady=5)

    def enviar_correo(self, email, contraseña):
        asunto = self.entry_asunto.get()
        cuerpo = self.textbox_cuerpo.get("1.0", "end-1c")
        from_email = self.entry_from.get() or email  # Usa el correo especificado o el correo de autenticación

        # Verificar si el DataFrame está vacío
        if self.data_frame is None or self.data_frame.empty:
            messagebox.showerror("Error", "Debe adjuntar un archivo Excel válido.")
            return

        try:
            smtp_server = 'smtp.gmail.com'
            smtp_port = 587
            smtp_username = email
            smtp_password = contraseña

            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(smtp_username, smtp_password)

            for index, row in self.data_frame.iterrows():
                correos = row.get('CORREOS')
                if isinstance(correos, str):
                    emails = [correo.strip() for correo in correos.split(';') if correo.strip()]
                    codigo_compania = row.get('ID', '')
                    nombre_compañia = row.get('NOMBRE', '')

                    if emails:
                        msg = MIMEMultipart()
                        msg['From'] = from_email
                        msg['To'] = ', '.join(emails)  # Todos los correos en un solo campo
                        msg['Subject'] = f'{asunto}'
                        msg.attach(MIMEText(cuerpo, 'plain'))

                        # Adjuntar archivo común solo si ha sido seleccionado
                        if self.archivo_a_enviar:
                            with open(self.archivo_a_enviar, 'rb') as file:
                                attachment = MIMEApplication(file.read(), Name=self.archivo_a_enviar)
                                attachment.add_header('Content-Disposition', 'attachment', filename=self.archivo_a_enviar.split("/")[-1])
                                msg.attach(attachment)

                        # Enviar el correo a todos los destinatarios
                        server.sendmail(from_email, emails, msg.as_string())

                        # Mostrar mensaje en la app para el envío de correos
                        mensaje = f"Correo enviado a: {', '.join(emails)}"
                        print(mensaje)
                        ctk.CTkLabel(self.nueva_ventana, text=mensaje, fg_color="green").pack(pady=5)

            server.quit()

            # Mostrar mensaje de proceso finalizado
            print("Proceso terminado.")
            ctk.CTkLabel(self.nueva_ventana, text="Proceso finalizado", fg_color="blue").pack(pady=10)
            messagebox.showinfo("Éxito", "Correos enviados exitosamente.")
        except Exception as e:
            print(f"Error al enviar correos: {e}")
            messagebox.showerror("Error", f"Error al enviar correos: {e}")

if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    app = App(root)
    root.mainloop()

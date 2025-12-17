from selenium import webdriver
from selenium.webdriver.common.by import By
import tkinter as tk
from tkinter import messagebox
import time
from threading import Thread
from tkinter import ttk
import os
from tkinter import filedialog
import pymupdf
from datetime import datetime


class App(tk.Tk):
    def __init__(self, screenName = None, baseName = None, className = "Tk", useTk = True, sync = False, use = None):
        super().__init__(screenName, baseName, className, useTk, sync, use)
        self.title("Super Descarga Masiva T-Registro")
        self.geometry("400x110")

        self.columnconfigure(0,minsize=124)
        self.columnconfigure(1,minsize=124)
        self.columnconfigure(2,minsize=124)

        self.driver = None

        self.label_nro_pagina = tk.Label(self,text="Nro Página Inicio")
        self.label_nro_pagina.grid(column=0,row=0)

        self.entry_nro_pagina = tk.Entry(self)
        self.entry_nro_pagina.grid(column=0,row=1,padx=5)

        self.label_nro_posicion = tk.Label(self,text="Nro Documento Inicio")
        self.label_nro_posicion.grid(column=1,row=0)

        self.entry_nro_documento = tk.Entry(self)
        self.entry_nro_documento.grid(column=1,row=1,padx=5)

        self.btn_abrir_chrome = tk.Button(self,text="Abrir Chrome",command=self.on_btn_abrir_chrome_pressed)
        self.btn_abrir_chrome.grid(column=2,row=0,sticky="we",pady=1,padx=5)

        self.btn_descarga_cirs = tk.Button(self,text="Iniciar Descarga",command=self.on_btn_descargar_cirs_pressed)
        self.btn_descarga_cirs.grid(column=2,row=1,sticky="we",pady=1,padx=5)

        self.progressbar_descargas = ttk.Progressbar(self,mode="indeterminate")
        # self.progressbar_descargas.grid(row=2,column=0,columnspan=2)

        self.btn_renombrar_cirs = tk.Button(self,text="Renombrar PDF CIR",command=self.on_renombrar_pdfs_cir_pressed)
        self.btn_renombrar_cirs.grid(row=2,column=2,sticky="we",pady=1,padx=5)

        self.label_autor = tk.Label(self, text="Hecho por José Melgarejo. Difunde y comparte!!!") ## NO ME BORRES LOCO ##
        self.label_autor.grid(row=3,column=0,columnspan=3,sticky="we") ## NO ME BORRES LOCO ##
        
    def on_renombrar_pdfs_cir_pressed(self):
        opcion = messagebox.askquestion(
        "Seleccionar origen",
        "¿Qué deseas seleccionar?",
        icon="question",
        type=messagebox.YESNO,
        default=messagebox.YES,
        detail="Sí = Carpeta\nNo = PDF(s)")

        if opcion == "yes":
            carpeta = filedialog.askdirectory(title="Selecciona una carpeta")
            if not carpeta:
                return
            pdfs = [os.path.join(carpeta,file) for file in os.listdir(carpeta) if file.lower().endswith(".pdf")]
            for pdf in pdfs:
                self.renombrar_pdf_cir(pdf)
            messagebox.showinfo(message="Renombrado Finalizado :)")

        else:
            archivos = filedialog.askopenfilenames(
            title="Selecciona uno o varios PDFs",
            filetypes=[("Archivos PDF", "*.pdf")])

            if not archivos:
                return

            for pdf in archivos:
                self.renombrar_pdf_cir(pdf)
            messagebox.showinfo(message="Renombrado Finalizado :)")

    def renombrar_pdf_cir(self, pdf):
        try:
            now = datetime.now().strftime("%Y%m%d%H%M%S")
            doc = pymupdf.open(pdf)
            page = doc[0]
            pagetext = page.get_text(sort=True).split("\n")
            dni = next((line.split("-")[1][:10].strip() for line in pagetext if "Tipo y número de documento" in line), None)
            nombres = next((line.split(":")[1].replace(",", "").strip() for line in pagetext if "Apellidos y nombres" in line), None)
            tipo_pdf = pagetext[3].strip()[14:18]
            folder_input = os.path.dirname(pdf)
            
            new_filename = os.path.join(folder_input,f"{dni} {nombres} {tipo_pdf} {now}.pdf")
            doc.save(new_filename)
        except:
            pass

    def on_btn_abrir_chrome_pressed(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--log-level=1")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        prefs = {
            "profile.default_content_settings.popups": 0,    
            # "download.default_directory": os.path.abspath(self.download_folder_path), # Cambiar la ruta de descargas
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,   # <--- PDF se descarga, no se abre
            "profile.default_content_setting_values.automatic_downloads": 1
        }
        options.add_experimental_option("prefs", prefs)
        options.add_experimental_option("detach", True)
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--disable-infobars")

        self.driver = webdriver.Chrome(options=options)
        self.driver.implicitly_wait(10)
        self.driver.get(r"https://api-seguridad.sunat.gob.pe/v1/clientessol/4f3b88b3-d9d6-402a-b85d-6a0bc857746a/oauth2/loginMenuSol?lang=es-PE&showDni=true&showLanguages=false&originalUrl=https://e-menu.sunat.gob.pe/cl-ti-itmenu/AutenticaMenuInternet.htm&state=rO0ABXNyABFqYXZhLnV0aWwuSGFzaE1hcAUH2sHDFmDRAwACRgAKbG9hZEZhY3RvckkACXRocmVzaG9sZHhwP0AAAAAAAAx3CAAAABAAAAADdAAEZXhlY3B0AAZwYXJhbXN0AEsqJiomL2NsLXRpLWl0bWVudS9NZW51SW50ZXJuZXQuaHRtJmI2NGQyNmE4YjVhZjA5MTkyM2IyM2I2NDA3YTFjMWRiNDFlNzMzYTZ0AANleGVweA==")


    def on_btn_descargar_cirs_pressed(self):
        new_thread = Thread(
            target= self.iniciar_descarga_cirs,
            daemon=True
        )
        new_thread.start()


    def iniciar_descarga_cirs(self):
        if self.driver:
            try: 
                self.progressbar_descargas.grid(row=2,column=0,columnspan=2,sticky="we")
                self.progressbar_descargas.start()
                pagina_inicio = self.entry_nro_pagina.get()
                nro_documento_inicio = self.entry_nro_documento.get()
                self.driver.switch_to.frame("iframeApplication")
                self.success_downloads = 0
                self.last_documento = None

                header = self.driver.find_element(By.CLASS_NAME, "headPagePages")
                pages = header.find_elements(By.TAG_NAME, "a")
                num_pages = [p.text for p in pages]

                def click_and_download(cir_list):
                    # cir_list = self.driver.find_elements(By.XPATH, '//a[starts-with(@href, "javascript:descargarCIR")]')
                    for l in cir_list:
                        time.sleep(1)
                        l.click()
                        self.success_downloads += 1
                        self.last_documento = str(l.get_attribute("href")).split("|")[3]

                if num_pages:
                    for page in range(int(pagina_inicio),int(num_pages[-1])+1):
                        try:
                            selected_page = self.driver.find_element(By.XPATH, f"//a[starts-with(@href, 'javascript:') and starts-with(@class, 'page') and text() = '{page}']")
                            selected_page.click()
                            cir_list = self.driver.find_elements(By.XPATH, '//a[starts-with(@href, "javascript:descargarCIR")]')
                            if page == int(pagina_inicio) and nro_documento_inicio:
                                for i, elem in enumerate(cir_list):
                                    href = elem.get_attribute("href")
                                    if str(nro_documento_inicio) in href:
                                        cir_list = cir_list[i:]
                                        break
                            click_and_download(cir_list)
                        except Exception as e:
                            messagebox.showerror(message=f"Error al procesar la página {page}. EL último Documento es {self.last_documento}.")
                            break
                else:
                    click_and_download()
                self.progressbar_descargas.stop()
                messagebox.showinfo(message= f"Ha culminado al descarga de {self.success_downloads} archivos CIR.")
            
            except Exception as e:
                messagebox.showerror(title="Error", message=f"Ocurrió un error durante la descarga de CIR: {e}")

            finally:
                self.driver.switch_to.default_content()

                self.progressbar_descargas.grid_remove()
        else:
            messagebox.showwarning(message="Debe Abrir Chrome primero.")

if __name__ == "__main__":
    app = App()
    app.mainloop()
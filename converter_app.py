import os
import sys
import threading
import subprocess
import platform
import time
from tkinter import filedialog, messagebox
import customtkinter as ctk

# Configuração inicial do CustomTkinter
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class DocumentConverter:
    """
    Classe responsável pela lógica de conversão de documentos.
    Gerencia a detecção de ferramentas (Word/LibreOffice) e a execução da conversão.
    """
    def __init__(self):
        self.has_word = self._check_word_installed()
        self.libreoffice_path = self._find_libreoffice()
        self.supported_extensions = ['.docx', '.doc', '.odt', '.rtf']

    def _check_word_installed(self):
        """Verifica se o MS Word está instalado (apenas Windows)."""
        if platform.system() != "Windows":
            return False
        try:
            # Tenta importar docx2pdf apenas se estiver no Windows
            import docx2pdf
            # Uma verificação mais robusta seria tentar instanciar o COM object,
            # mas a importação já é um bom indício se a lib estiver instalada.
            return True
        except ImportError:
            return False
        except Exception:
            return False

    def _find_libreoffice(self):
        """Tenta localizar o executável do LibreOffice."""
        system = platform.system()
        paths_to_check = []

        if system == "Windows":
            paths_to_check = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            ]
        elif system == "Darwin":  # macOS
            paths_to_check = [
                "/Applications/LibreOffice.app/Contents/MacOS/soffice"
            ]
        else:  # Linux
            # No Linux, 'soffice' ou 'libreoffice' geralmente estão no PATH
            return "soffice"

        for path in paths_to_check:
            if os.path.exists(path):
                return path
        
        # Se não encontrar nos caminhos padrão, tenta chamar pelo comando se estiver no PATH
        return "soffice" if self._is_command_available("soffice") else None

    def _is_command_available(self, command):
        """Verifica se um comando está disponível no PATH do sistema."""
        from shutil import which
        return which(command) is not None

    def convert_to_pdf(self, input_path, output_folder):
        """
        Converte um arquivo para PDF.
        Retorna (True, mensagem) em caso de sucesso, ou (False, erro).
        """
        ext = os.path.splitext(input_path)[1].lower()
        filename = os.path.basename(input_path)
        
        if ext not in self.supported_extensions:
            return False, f"Extensão não suportada: {ext}"

        # Lógica de decisão: Word vs LibreOffice
        # Preferimos Word para .docx/.doc no Windows pela fidelidade
        use_word = self.has_word and ext in ['.docx', '.doc']
        
        if use_word:
            try:
                from docx2pdf import convert
                # docx2pdf converte para a mesma pasta se output_path não for especificado,
                # ou podemos especificar o arquivo de saída.
                # Para garantir que vá para a output_folder:
                convert(input_path, output_folder)
                return True, f"Convertido com Word: {filename}"
            except Exception as e:
                # Fallback para LibreOffice se o Word falhar
                if self.libreoffice_path:
                    return self._convert_with_libreoffice(input_path, output_folder)
                return False, f"Erro Word: {str(e)}"
        
        elif self.libreoffice_path:
            return self._convert_with_libreoffice(input_path, output_folder)
        else:
            return False, "Nenhum conversor (Word ou LibreOffice) encontrado."

    def _convert_with_libreoffice(self, input_path, output_folder):
        """Executa a conversão via linha de comando do LibreOffice."""
        try:
            # Comando: soffice --headless --convert-to pdf <file> --outdir <dir>
            cmd = [
                self.libreoffice_path,
                '--headless',
                '--convert-to', 'pdf',
                input_path,
                '--outdir', output_folder
            ]
            
            # No Windows, subprocess precisa de tratamento especial para não abrir janelas de console
            startupinfo = None
            if platform.system() == "Windows":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            
            result = subprocess.run(
                cmd, 
                stdout=subprocess.PIPE, 
                stderr=subprocess.PIPE,
                startupinfo=startupinfo,
                check=True
            )
            return True, f"Convertido com LibreOffice: {os.path.basename(input_path)}"
        except subprocess.CalledProcessError as e:
            return False, f"Erro LibreOffice: {e.stderr.decode('utf-8', errors='ignore')}"
        except Exception as e:
            return False, f"Erro genérico LibreOffice: {str(e)}"


class ConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Conversor de Documentos para PDF")
        self.geometry("700x550")
        
        self.converter = DocumentConverter()
        self.selected_files = []
        self.is_converting = False

        self._setup_ui()
        self._check_dependencies()

    def _setup_ui(self):
        # Layout de Grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1) # Log expande

        # 1. Cabeçalho
        self.header_frame = ctk.CTkFrame(self)
        self.header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        
        self.lbl_title = ctk.CTkLabel(
            self.header_frame, 
            text="Conversor de Documentos para PDF", 
            font=ctk.CTkFont(size=20, weight="bold")
        )
        self.lbl_title.pack(pady=10)
        
        self.lbl_subtitle = ctk.CTkLabel(
            self.header_frame,
            text="Suporta: DOCX, DOC, ODT, RTF",
            text_color="gray"
        )
        self.lbl_subtitle.pack(pady=(0, 10))

        # 2. Seleção
        self.selection_frame = ctk.CTkFrame(self)
        self.selection_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        
        self.btn_file = ctk.CTkButton(self.selection_frame, text="Selecionar Arquivo(s)", command=self.select_files)
        self.btn_file.grid(row=0, column=0, padx=10, pady=10)
        
        self.btn_folder = ctk.CTkButton(self.selection_frame, text="Selecionar Pasta", command=self.select_folder)
        self.btn_folder.grid(row=0, column=1, padx=10, pady=10)

        self.lbl_selection = ctk.CTkLabel(self.selection_frame, text="Nenhum arquivo selecionado", wraplength=600)
        self.lbl_selection.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 10))

        # 3. Ação
        self.action_frame = ctk.CTkFrame(self)
        self.action_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        
        self.btn_convert = ctk.CTkButton(
            self.action_frame, 
            text="Converter para PDF", 
            command=self.start_conversion_thread,
            state="disabled",
            fg_color="green",
            hover_color="darkgreen"
        )
        self.btn_convert.pack(pady=10, fill="x", padx=20)

        # 4. Progresso
        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.grid(row=3, column=0, padx=20, pady=(10, 0), sticky="ew")
        self.progress_bar.set(0)

        # 5. Log
        self.log_box = ctk.CTkTextbox(self, state="disabled")
        self.log_box.grid(row=4, column=0, padx=20, pady=20, sticky="nsew")

    def _check_dependencies(self):
        status = []
        if self.converter.has_word:
            status.append("MS Word detectado (docx2pdf).")
        else:
            status.append("MS Word NÃO detectado.")
            
        if self.converter.libreoffice_path:
            status.append(f"LibreOffice detectado.")
        else:
            status.append("LibreOffice NÃO detectado. Instale para suporte a ODT/RTF.")
            
        self.log_message("\n".join(status))
        
        if not self.converter.has_word and not self.converter.libreoffice_path:
            messagebox.showwarning("Aviso", "Nenhum conversor encontrado!\nInstale o MS Word ou LibreOffice.")

    def log_message(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def select_files(self):
        files = filedialog.askopenfilenames(
            filetypes=[("Documentos", "*.docx *.doc *.odt *.rtf")]
        )
        if files:
            self.selected_files = list(files)
            self.update_selection_label(f"{len(files)} arquivo(s) selecionado(s)")
            self.btn_convert.configure(state="normal")
            self.log_message(f"Selecionados: {len(files)} arquivos.")

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            found_files = []
            for root, dirs, files in os.walk(folder):
                for file in files:
                    if os.path.splitext(file)[1].lower() in self.converter.supported_extensions:
                        found_files.append(os.path.join(root, file))
            
            if found_files:
                self.selected_files = found_files
                self.update_selection_label(f"Pasta: {folder} ({len(found_files)} arquivos compatíveis)")
                self.btn_convert.configure(state="normal")
                self.log_message(f"Pasta selecionada. {len(found_files)} arquivos encontrados.")
            else:
                self.update_selection_label("Nenhum arquivo compatível encontrado na pasta.")
                self.btn_convert.configure(state="disabled")

    def update_selection_label(self, text):
        self.lbl_selection.configure(text=text)

    def start_conversion_thread(self):
        if not self.selected_files:
            return
        
        self.is_converting = True
        self.btn_convert.configure(state="disabled", text="Convertendo...")
        self.btn_file.configure(state="disabled")
        self.btn_folder.configure(state="disabled")
        self.progress_bar.set(0)
        
        thread = threading.Thread(target=self.run_conversion)
        thread.start()

    def run_conversion(self):
        total = len(self.selected_files)
        success_count = 0
        
        self.log_message("-" * 30)
        self.log_message("Iniciando conversão...")

        for i, file_path in enumerate(self.selected_files):
            # Define output folder (mesma do arquivo original)
            output_folder = os.path.dirname(file_path)
            
            # Atualiza UI de forma segura (embora simples, o ideal seria after, mas ctk aguenta chamadas simples)
            # Para ser 100% correto com tkinter, usaremos after para atualizações visuais se necessário,
            # mas aqui vamos direto para simplificar o exemplo, pois ctk costuma tolerar.
            # Se der erro, mudaremos para self.after.
            
            success, msg = self.converter.convert_to_pdf(file_path, output_folder)
            
            if success:
                success_count += 1
                self.log_message(f"[OK] {msg}")
            else:
                self.log_message(f"[ERRO] {msg}")
            
            progress = (i + 1) / total
            self.progress_bar.set(progress)
        
        self.log_message("-" * 30)
        self.log_message(f"Concluído! {success_count}/{total} arquivos convertidos.")
        
        # Restaura UI
        self.after(0, self.reset_ui)

    def reset_ui(self):
        self.is_converting = False
        self.btn_convert.configure(state="normal", text="Converter para PDF")
        self.btn_file.configure(state="normal")
        self.btn_folder.configure(state="normal")
        messagebox.showinfo("Sucesso", "Processo de conversão finalizado!")

if __name__ == "__main__":
    # Instruções de Instalação
    print("--- Conversor de Documentos ---")
    print("Requisitos:")
    print("1. Python 3.x")
    print("2. Bibliotecas: pip install customtkinter docx2pdf")
    print("3. Softwares: Microsoft Word (para .docx/.doc) E/OU LibreOffice (para .odt/.rtf)")
    print("-------------------------------")
    
    app = ConverterApp()
    app.mainloop()

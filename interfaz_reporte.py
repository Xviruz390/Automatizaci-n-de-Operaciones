"""Interfaz sencilla para ejecutar la automatización de reportes."""
from __future__ import annotations

import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

ROOT = Path(__file__).resolve().parent
SCRIPT = ROOT / "automatizar_reporte.py"


class ReporteApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Automatización de reportes")
        self.geometry("620x390")
        self.resizable(False, False)
        self.process: subprocess.Popen[str] | None = None
        self.date_var = tk.StringVar(value=datetime.now().strftime("%d-%m-%Y"))
        self.status_var = tk.StringVar(value="Listo para ejecutar")
        self._build_ui()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self, padding=24)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Automatización de Reporte 2", font=("Segoe UI", 16, "bold")).pack(anchor="w")
        ttk.Label(
            frame,
            text="Seleccione la fecha de las operaciones del Reporte 1.",
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(6, 20))

        date_frame = ttk.Frame(frame)
        date_frame.pack(anchor="w")
        ttk.Label(date_frame, text="Fecha (DD-MM-AAAA):").pack(side="left")
        self.date_entry = ttk.Entry(date_frame, textvariable=self.date_var, width=16)
        self.date_entry.pack(side="left", padx=(12, 0))
        self.date_entry.focus_set()

        buttons = ttk.Frame(frame)
        buttons.pack(anchor="w", pady=18)
        self.run_button = ttk.Button(buttons, text="Ejecutar automatización", command=self.run_automation)
        self.run_button.pack(side="left")
        ttk.Button(buttons, text="Abrir carpeta", command=self.open_folder).pack(side="left", padx=10)

        ttk.Label(frame, textvariable=self.status_var, foreground="#555").pack(anchor="w")
        ttk.Label(frame, text="Resultado de la ejecución:", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(20, 5))
        self.output = tk.Text(frame, height=8, width=70, state="disabled", wrap="word")
        self.output.pack(fill="both", expand=True)

    def run_automation(self) -> None:
        try:
            selected_date = datetime.strptime(self.date_var.get().strip(), "%d-%m-%Y")
        except ValueError:
            messagebox.showerror("Fecha inválida", "Ingrese la fecha con el formato DD-MM-AAAA.")
            return

        self.run_button.configure(state="disabled")
        self.status_var.set("Ejecutando...")
        self._write_output(f"Procesando {selected_date:%d-%m-%Y}...\n")
        thread = threading.Thread(target=self._run_process, args=(selected_date,), daemon=True)
        thread.start()

    def _run_process(self, selected_date: datetime) -> None:
        command = [sys.executable, str(SCRIPT), "--fecha", selected_date.strftime("%Y-%m-%d")]
        try:
            self.process = subprocess.Popen(
                command,
                cwd=ROOT,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            output, _ = self.process.communicate()
            code = self.process.returncode
        except Exception as error:
            output, code = f"Error al ejecutar: {error}\n", 1
        self.after(0, self._finish_process, output, code)

    def _finish_process(self, output: str, code: int) -> None:
        self._write_output(output)
        self.status_var.set("Proceso terminado correctamente" if code == 0 else "El proceso terminó con errores")
        self.run_button.configure(state="normal")
        if code == 0:
            messagebox.showinfo("Listo", "El reporte fue generado correctamente.")
        else:
            messagebox.showerror("Error", "No se pudo generar el reporte. Revise el resultado mostrado.")

    def _write_output(self, text: str) -> None:
        self.output.configure(state="normal")
        self.output.delete("1.0", "end")
        self.output.insert("end", text)
        self.output.configure(state="disabled")

    def open_folder(self) -> None:
        subprocess.Popen(["explorer", str(ROOT)])


if __name__ == "__main__":
    ReporteApp().mainloop()

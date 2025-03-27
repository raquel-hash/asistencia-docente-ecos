import tkinter as tk
from registroDocente import abrir_lista_docentes
from materias import abrir_lista_materias
from horarios import abrir_gestion_horarios
from asistencia import abrir_registro_asistencia
from reporte_mensual import abrir_reporte

def main():
    """Inicializa la aplicación principal."""
    root_menu = tk.Tk()
    root_menu.title("Sistema de Gestión de Docentes")
    root_menu.geometry("400x200")
    root_menu.state('zoomed')  # Ocupa toda la pantalla
    root_menu.configure(bg='#e0e0e0')

    tk.Label(root_menu, text="Sistema de Gestión de Docentes", font=("Arial", 16), bg='#e0e0e0').pack(pady=20)

    # Botón para abrir la lista de docentes
    tk.Button(root_menu, text="Lista de Docentes", command=lambda: abrir_lista_docentes(root_menu), 
              bg='#d32f2f', fg='white', width=20).pack(pady=10)
    
    # Botón para abrir la lista de materias
    tk.Button(root_menu, text="Lista de Materias", command=lambda: abrir_lista_materias(root_menu), 
              bg='#d32f2f', fg='white', width=20).pack(pady=10)

    # Botón para abrir la gestión de horarios
    tk.Button(root_menu, text="Gestión de Horarios", command=lambda: abrir_gestion_horarios(root_menu), 
              bg='#d32f2f', fg='white', width=20).pack(pady=10)

    # Botón para abrir el registro de asistencia
    tk.Button(root_menu, text="Registro de Asistencia", command=lambda: abrir_registro_asistencia(root_menu), 
              bg='#d32f2f', fg='white', width=20).pack(pady=10)
    
    # Botón para abrir el reporte
    tk.Button(root_menu, text="Generar Reporte", command=lambda: abrir_reporte(root_menu), 
              bg='#d32f2f', fg='white', width=20).pack(pady=10)

    # Botón para salir de la aplicación
    tk.Button(root_menu, text="Salir", command=root_menu.quit, bg='#b71c1c', fg='white', width=20).pack(pady=10)

    root_menu.mainloop()

if __name__ == "__main__":
    main()

import sys
import subprocess
import importlib
import tkinter as tk
from tkinter import messagebox
import threading
import os

# Global flag to prevent multiple installation windows
installing = False

def check_and_install_modules():
    """Check and install required modules"""
    global installing
    
    if installing:
        return  # Prevent multiple installation windows
    
    required_modules = [
        ('pandas', 'pandas'),
        ('docxtpl', 'python-docx-template'),
        ('docx', 'python-docx')
        # Removed pywin32 as it's bundled with the exe
    ]
    
    missing_modules = []
    
    # Check which modules are missing
    for module_name, pip_name in required_modules:
        try:
            if module_name == 'docx':
                # Special check for python-docx
                import docx
            else:
                importlib.import_module(module_name)
        except ImportError:
            missing_modules.append((module_name, pip_name))
    
    # Special check for win32com (part of pywin32)
    try:
        import win32com.client
    except ImportError:
        # Only add pywin32 if win32com is actually missing
        missing_modules.append(('win32com', 'pywin32'))
    
    if missing_modules:
        installing = True
        
        # Show installation dialog
        root = tk.Tk()
        root.withdraw()
        root.lift()
        root.attributes('-topmost', True)
        
        modules_list = ', '.join([m[0] for m in missing_modules])
        result = messagebox.askyesno(
            "Missing Dependencies", 
            f"The following modules are missing: {modules_list}\n\n"
            "Do you want to install them automatically?\n"
            "This may take a few minutes.",
            parent=root
        )
        
        if result:
            # Show progress window
            progress_window = tk.Toplevel(root)
            progress_window.title("Installing Dependencies")
            progress_window.geometry("450x200")
            progress_window.resizable(False, False)
            progress_window.grab_set()  # Make it modal
            progress_window.lift()
            progress_window.attributes('-topmost', True)
            
            # Center the window
            progress_window.update_idletasks()
            x = (progress_window.winfo_screenwidth() // 2) - (450 // 2)
            y = (progress_window.winfo_screenheight() // 2) - (200 // 2)
            progress_window.geometry(f"450x200+{x}+{y}")
            
            status_label = tk.Label(progress_window, text="Installing modules...", 
                                  font=('Segoe UI', 12, 'bold'))
            status_label.pack(pady=10)
            
            progress_text = tk.Text(progress_window, height=8, width=55, font=('Consolas', 9))
            progress_text.pack(pady=10, padx=20)
            
            def install_modules():
                try:
                    for module_name, pip_name in missing_modules:
                        progress_text.insert(tk.END, f"Installing {module_name}...\n")
                        progress_text.see(tk.END)
                        progress_window.update()
                        
                        result = subprocess.run([sys.executable, '-m', 'pip', 'install', pip_name], 
                                              capture_output=True, text=True)
                        
                        if result.returncode == 0:
                            progress_text.insert(tk.END, f"✓ {module_name} installed successfully\n")
                        else:
                            progress_text.insert(tk.END, f"✗ Failed to install {module_name}\n")
                            progress_text.insert(tk.END, f"Error: {result.stderr[:100]}...\n")
                        
                        progress_text.see(tk.END)
                        progress_window.update()
                    
                    progress_text.insert(tk.END, "\nInstallation complete! Starting application...")
                    progress_text.see(tk.END)
                    progress_window.update()
                    
                    # Wait a moment then close and start app
                    progress_window.after(2000, lambda: [progress_window.destroy(), root.destroy(), start_app()])
                    
                except Exception as e:
                    progress_text.insert(tk.END, f"\nInstallation failed: {str(e)}")
                    progress_window.after(3000, lambda: [progress_window.destroy(), root.destroy()])
            
            # Start installation in thread
            threading.Thread(target=install_modules, daemon=True).start()
            root.mainloop()
        else:
            root.destroy()
            messagebox.showwarning("Cannot Start", "Application requires these modules to run.")
            sys.exit(1)
    else:
        start_app()

def start_app():
    """Start the main application"""
    try:
        from app import SaralWorksApp
        app = SaralWorksApp()
        app.run()
    except Exception as e:
        messagebox.showerror("Application Error", f"Failed to start application: {str(e)}")

if __name__ == "__main__":
    check_and_install_modules()
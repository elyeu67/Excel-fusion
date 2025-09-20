import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from pathlib import Path
import threading
from datetime import datetime
import sys
import ctypes
from ctypes import wintypes

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        
        # Optimisations pour la qualité visuelle
        self.setup_high_dpi_support()
        self.setup_visual_quality()
        
        self.root.title("📊 Fusionneur de Fichiers Excel - Interface Moderne")
        self.root.geometry("900x750")
        self.root.resizable(True, True)
        
        # Configuration de la fenêtre moderne
        self.root.configure(bg='#f8f9fa')
        self.root.minsize(800, 650)
        
        # Variables
        self.input_folder = tk.StringVar()
        self.output_file = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="🚀 Prêt à fusionner des fichiers Excel")
        
        # Couleurs vives et dynamiques
        self.colors = {
            'primary': '#6366f1',      # Indigo vif
            'secondary': '#8b5cf6',    # Violet
            'success': '#10b981',      # Vert émeraude
            'danger': '#ef4444',       # Rouge vif
            'warning': '#f59e0b',      # Orange ambré
            'info': '#06b6d4',         # Cyan
            'light': '#f1f5f9',        # Gris très clair
            'dark': '#1e293b',         # Bleu foncé
            'white': '#ffffff',
            'border': '#e2e8f0',
            'text': '#0f172a',
            'text_muted': '#64748b',
            'accent1': '#ec4899',      # Rose vif
            'accent2': '#84cc16',      # Vert lime
            'accent3': '#f97316',      # Orange vif
            'gradient_start': '#667eea',
            'gradient_end': '#764ba2'
        }
        
        # Variables pour les animations
        self.animation_running = False
        self.animation_step = 0
        
        self.setup_modern_ui()
        self.start_background_animation()
    
    def start_background_animation(self):
        """Démarre l'animation de background colorée"""
        self.animation_running = True
        self.animate_background()
    
    def animate_background(self):
        """Animation du background avec des couleurs changeantes"""
        if not self.animation_running:
            return
            
        # Cycle à travers différentes couleurs
        colors_cycle = [
            self.colors['light'],
            '#fef3c7',  # Jaune très clair
            '#fce7f3',  # Rose très clair
            '#e0e7ff',  # Bleu très clair
            '#ecfdf5',  # Vert très clair
            '#f0f9ff',  # Cyan très clair
        ]
        
        current_color = colors_cycle[self.animation_step % len(colors_cycle)]
        self.root.configure(bg=current_color)
        
        # Mettre à jour tous les frames avec la nouvelle couleur
        self.update_frame_colors(current_color)
        
        self.animation_step += 1
        
        # Programmer la prochaine animation (changement toutes les 3 secondes)
        self.root.after(3000, self.animate_background)
    
    def update_frame_colors(self, bg_color):
        """Met à jour les couleurs des frames pour l'animation"""
        try:
            # Mettre à jour les frames principaux
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Frame) and widget.cget('bg') == self.colors['light']:
                    widget.configure(bg=bg_color)
        except:
            pass  # Ignorer les erreurs de widgets supprimés
    
    def setup_high_dpi_support(self):
        """Configure le support DPI élevé pour Windows"""
        try:
            # Activer le support DPI élevé sur Windows
            if sys.platform == "win32":
                # Définir le DPI awareness
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
                
                # Obtenir le facteur de mise à l'échelle DPI
                user32 = ctypes.windll.user32
                user32.SetProcessDPIAware()
                
                # Configurer tkinter pour le DPI élevé
                self.root.tk.call('tk', 'scaling', 1.0)
                
        except Exception as e:
            print(f"Note: Impossible de configurer le DPI élevé: {e}")
    
    def setup_visual_quality(self):
        """Configure la qualité visuelle de l'interface"""
        try:
            # Configuration pour améliorer la qualité des polices
            self.root.tk.call('tk', 'fontconfigure', 'TkDefaultFont', '-family', 'Segoe UI')
            self.root.tk.call('tk', 'fontconfigure', 'TkTextFont', '-family', 'Segoe UI')
            self.root.tk.call('tk', 'fontconfigure', 'TkFixedFont', '-family', 'Consolas')
            
            # Améliorer l'anti-aliasing des polices
            if sys.platform == "win32":
                # Configuration pour Windows
                self.root.tk.call('tk', 'fontconfigure', 'TkDefaultFont', '-size', 9)
                self.root.tk.call('tk', 'fontconfigure', 'TkTextFont', '-size', 9)
                self.root.tk.call('tk', 'fontconfigure', 'TkFixedFont', '-size', 9)
            
            # Configuration de la qualité d'affichage
            self.root.tk.call('tk', 'scaling', 1.25)  # Augmenter légèrement l'échelle
            
        except Exception as e:
            print(f"Note: Impossible de configurer la qualité visuelle: {e}")
        
    def setup_modern_ui(self):
        # Configuration des styles modernes
        self.setup_modern_styles()
        
        # Frame principal avec padding moderne
        main_frame = tk.Frame(self.root, bg=self.colors['light'], padx=30, pady=25)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Configuration du grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Header moderne avec gradient effect
        self.create_modern_header(main_frame)
        
        # Section de sélection des fichiers avec design moderne
        self.create_file_selection_section(main_frame)
        
        # Options avec design moderne
        self.create_options_section(main_frame)
        
        # Bouton d'action principal
        self.create_action_button(main_frame)
        
        # Barre de progression moderne
        self.create_progress_section(main_frame)
        
        # Zone de logs moderne
        self.create_logs_section(main_frame)
        
    def setup_modern_styles(self):
        """Configure les styles modernes pour l'interface avec qualité optimisée"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configuration des polices avec qualité optimisée
        font_family = 'Segoe UI'
        font_size_large = 11
        font_size_medium = 10
        font_size_small = 9
        
        # Style pour les boutons modernes avec couleurs vives
        style.configure('Modern.TButton',
                       background=self.colors['primary'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none',
                       font=(font_family, font_size_medium, 'bold'),
                       padding=(25, 14))
        
        style.map('Modern.TButton',
                 background=[('active', self.colors['accent1']),
                           ('pressed', self.colors['accent3']),
                           ('disabled', '#94a3b8')])
        
        # Style pour les boutons secondaires avec couleurs vives
        style.configure('Secondary.TButton',
                       background=self.colors['secondary'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none',
                       font=(font_family, font_size_small, 'normal'),
                       padding=(18, 10))
        
        style.map('Secondary.TButton',
                 background=[('active', self.colors['accent2']),
                           ('pressed', self.colors['accent3']),
                           ('disabled', '#94a3b8')])
        
        # Style pour les boutons avec couleurs d'accent
        style.configure('Accent.TButton',
                       background=self.colors['accent1'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none',
                       font=(font_family, font_size_small, 'bold'),
                       padding=(15, 8))
        
        style.map('Accent.TButton',
                 background=[('active', self.colors['accent2']),
                           ('pressed', self.colors['accent3']),
                           ('disabled', '#94a3b8')])
        
        # Style pour les frames avec bordures modernes
        style.configure('Modern.TFrame',
                       background=self.colors['white'],
                       relief='flat',
                       borderwidth=1)
        
        # Style pour les labels modernes avec polices optimisées
        style.configure('Title.TLabel',
                       background=self.colors['light'],
                       foreground=self.colors['dark'],
                       font=(font_family, 26, 'bold'))
        
        style.configure('Subtitle.TLabel',
                       background=self.colors['light'],
                       foreground=self.colors['text_muted'],
                       font=(font_family, font_size_medium))
        
        style.configure('Section.TLabel',
                       background=self.colors['white'],
                       foreground=self.colors['dark'],
                       font=(font_family, font_size_large, 'bold'))
        
        # Style pour les barres de progression avec couleurs vives
        style.configure('Modern.Horizontal.TProgressbar',
                       background=self.colors['accent1'],
                       troughcolor=self.colors['border'],
                       borderwidth=0,
                       lightcolor=self.colors['accent2'],
                       darkcolor=self.colors['accent3'],
                       thickness=10)
        
        # Style pour les Entry avec qualité améliorée
        style.configure('Modern.TEntry',
                       fieldbackground=self.colors['light'],
                       foreground=self.colors['text'],
                       borderwidth=1,
                       relief='flat',
                       font=(font_family, font_size_medium))
        
        # Style pour les Scrollbar
        style.configure('Modern.Vertical.TScrollbar',
                       background=self.colors['border'],
                       troughcolor=self.colors['light'],
                       borderwidth=0,
                       arrowcolor=self.colors['text_muted'],
                       darkcolor=self.colors['border'],
                       lightcolor=self.colors['border'])
    
    def create_modern_header(self, parent):
        """Crée un header moderne avec titre et sous-titre"""
        header_frame = tk.Frame(parent, bg=self.colors['light'])
        header_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 30))
        
        # Titre principal avec icône et qualité optimisée
        title_label = tk.Label(header_frame, 
                              text="📊 Fusionneur de Fichiers Excel",
                              font=('Segoe UI', 30, 'bold'),
                              fg=self.colors['dark'],
                              bg=self.colors['light'])
        title_label.pack()
        
        # Sous-titre avec qualité améliorée
        subtitle_label = tk.Label(header_frame,
                                 text="Fusionnez facilement des dizaines de fichiers Excel en un seul clic",
                                 font=('Segoe UI', 13),
                                 fg=self.colors['text_muted'],
                                 bg=self.colors['light'])
        subtitle_label.pack(pady=(8, 0))
    
    def create_file_selection_section(self, parent):
        """Crée la section de sélection des fichiers avec design coloré"""
        # Frame pour la sélection des fichiers avec bordure colorée
        selection_frame = tk.Frame(parent, bg=self.colors['white'], relief='flat', bd=3, highlightbackground=self.colors['accent1'], highlightthickness=2)
        selection_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        selection_frame.columnconfigure(1, weight=1)
        
        # Titre de section avec couleur vive
        section_title = tk.Label(selection_frame,
                                text="📁 Sélection des fichiers",
                                font=('Segoe UI', 15, 'bold'),
                                fg=self.colors['accent1'],
                                bg=self.colors['white'])
        section_title.grid(row=0, column=0, columnspan=3, sticky=tk.W, padx=20, pady=(20, 15))
        
        # Section dossier d'entrée
        input_frame = tk.Frame(selection_frame, bg=self.colors['white'])
        input_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), padx=20, pady=(0, 15))
        input_frame.columnconfigure(1, weight=1)
        
        tk.Label(input_frame, text="📂 Dossier contenant les fichiers Excel:",
                font=('Segoe UI', 11, 'bold'),
                fg=self.colors['accent2'],
                bg=self.colors['white']).grid(row=0, column=0, sticky=tk.W, pady=(0, 6))
        
        self.input_entry = tk.Entry(input_frame, textvariable=self.input_folder,
                                   font=('Segoe UI', 11),
                                   relief='flat',
                                   bd=2,
                                   bg=self.colors['light'],
                                   fg=self.colors['text'],
                                   insertbackground=self.colors['text'])
        self.input_entry.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 6))
        
        self.browse_input_btn = ttk.Button(input_frame, text="📁 Parcourir",
                                          command=self.browse_input_folder,
                                          style='Secondary.TButton')
        self.browse_input_btn.grid(row=2, column=0, sticky=tk.W, pady=(0, 10))
        
        # Section fichier de sortie
        output_frame = tk.Frame(selection_frame, bg=self.colors['white'])
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), padx=20, pady=(0, 20))
        output_frame.columnconfigure(1, weight=1)
        
        tk.Label(output_frame, text="💾 Fichier de sortie:",
                font=('Segoe UI', 11, 'bold'),
                fg=self.colors['accent3'],
                bg=self.colors['white']).grid(row=0, column=0, sticky=tk.W, pady=(0, 6))
        
        self.output_entry = tk.Entry(output_frame, textvariable=self.output_file,
                                    font=('Segoe UI', 11),
                                    relief='flat',
                                    bd=2,
                                    bg=self.colors['light'],
                                    fg=self.colors['text'],
                                    insertbackground=self.colors['text'])
        self.output_entry.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 6))
        
        self.browse_output_btn = ttk.Button(output_frame, text="💾 Enregistrer sous",
                                           command=self.browse_output_file,
                                           style='Secondary.TButton')
        self.browse_output_btn.grid(row=2, column=0, sticky=tk.W)
    
    def create_options_section(self, parent):
        """Crée la section des options avec design coloré"""
        options_frame = tk.Frame(parent, bg=self.colors['white'], relief='flat', bd=3, highlightbackground=self.colors['accent2'], highlightthickness=2)
        options_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        options_frame.columnconfigure(0, weight=1)
        
        # Titre de section avec couleur vive
        section_title = tk.Label(options_frame,
                                text="⚙️ Options de fusion",
                                font=('Segoe UI', 15, 'bold'),
                                fg=self.colors['accent2'],
                                bg=self.colors['white'])
        section_title.grid(row=0, column=0, sticky=tk.W, padx=20, pady=(20, 15))
        
        # Options avec checkboxes modernes
        options_content = tk.Frame(options_frame, bg=self.colors['white'])
        options_content.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=20, pady=(0, 20))
        
        self.add_source_column = tk.BooleanVar(value=True)
        self.ignore_headers = tk.BooleanVar(value=False)
        
        # Checkbox 1
        cb1_frame = tk.Frame(options_content, bg=self.colors['white'])
        cb1_frame.pack(fill=tk.X, pady=(0, 10))
        
        cb1 = tk.Checkbutton(cb1_frame,
                            text="📋 Ajouter une colonne avec le nom du fichier source",
                            variable=self.add_source_column,
                            font=('Segoe UI', 11),
                            fg=self.colors['accent1'],
                            bg=self.colors['white'],
                            selectcolor=self.colors['accent1'],
                            activebackground=self.colors['white'],
                            activeforeground=self.colors['accent1'])
        cb1.pack(side=tk.LEFT)
        
        # Checkbox 2
        cb2_frame = tk.Frame(options_content, bg=self.colors['white'])
        cb2_frame.pack(fill=tk.X)
        
        cb2 = tk.Checkbutton(cb2_frame,
                            text="📊 Ignorer les en-têtes dans les fichiers sources (garder seulement le premier)",
                            variable=self.ignore_headers,
                            font=('Segoe UI', 11),
                            fg=self.colors['accent3'],
                            bg=self.colors['white'],
                            selectcolor=self.colors['accent3'],
                            activebackground=self.colors['white'],
                            activeforeground=self.colors['accent3'])
        cb2.pack(side=tk.LEFT)
    
    def create_action_button(self, parent):
        """Crée le bouton d'action principal moderne"""
        button_frame = tk.Frame(parent, bg=self.colors['light'])
        button_frame.grid(row=3, column=0, columnspan=3, pady=(0, 20))
        
        self.merge_button = ttk.Button(button_frame,
                                      text="🚀 Fusionner les fichiers",
                                      command=self.start_merge,
                                      style='Modern.TButton')
        self.merge_button.pack()
    
    def create_progress_section(self, parent):
        """Crée la section de progression moderne"""
        progress_frame = tk.Frame(parent, bg=self.colors['light'])
        progress_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        progress_frame.columnconfigure(0, weight=1)
        
        # Barre de progression moderne
        self.progress_bar = ttk.Progressbar(progress_frame,
                                           variable=self.progress_var,
                                           maximum=100,
                                           style='Modern.Horizontal.TProgressbar')
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Label de statut moderne avec qualité améliorée
        self.status_label = tk.Label(progress_frame,
                                    textvariable=self.status_var,
                                    font=('Segoe UI', 11),
                                    fg=self.colors['primary'],
                                    bg=self.colors['light'])
        self.status_label.grid(row=1, column=0, sticky=tk.W)
    
    def create_logs_section(self, parent):
        """Crée la section des logs colorée"""
        logs_frame = tk.Frame(parent, bg=self.colors['white'], relief='flat', bd=3, highlightbackground=self.colors['accent3'], highlightthickness=2)
        logs_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        logs_frame.columnconfigure(0, weight=1)
        logs_frame.rowconfigure(1, weight=1)
        parent.rowconfigure(5, weight=1)
        
        # Titre de section avec couleur vive
        section_title = tk.Label(logs_frame,
                                text="📝 Journal des opérations",
                                font=('Segoe UI', 15, 'bold'),
                                fg=self.colors['accent3'],
                                bg=self.colors['white'])
        section_title.grid(row=0, column=0, sticky=tk.W, padx=20, pady=(20, 15))
        
        # Zone de texte moderne
        text_frame = tk.Frame(logs_frame, bg=self.colors['white'])
        text_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=20, pady=(0, 20))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(text_frame,
                               height=8,
                               wrap=tk.WORD,
                               font=('Consolas', 10),
                               bg=self.colors['light'],
                               fg=self.colors['text'],
                               relief='flat',
                               bd=2,
                               padx=12,
                               pady=12,
                               insertbackground=self.colors['text'])
        
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.log_text.yview, style='Modern.Vertical.TScrollbar')
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
    def browse_input_folder(self):
        folder = filedialog.askdirectory(title="Sélectionner le dossier contenant les fichiers Excel")
        if folder:
            self.input_folder.set(folder)
            self.log_message(f"Dossier sélectionné: {folder}")
            
    def browse_output_file(self):
        file = filedialog.asksaveasfilename(
            title="Enregistrer le fichier fusionné sous",
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel", "*.xlsx"), ("Tous les fichiers", "*.*")]
        )
        if file:
            self.output_file.set(file)
            self.log_message(f"Fichier de sortie: {file}")
            
    def log_message(self, message, level="info"):
        """Ajoute un message au journal avec un style moderne"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Couleurs selon le niveau
        colors = {
            "info": self.colors['primary'],
            "success": self.colors['success'],
            "warning": self.colors['warning'],
            "error": self.colors['danger']
        }
        
        # Icônes selon le niveau
        icons = {
            "info": "ℹ️",
            "success": "✅",
            "warning": "⚠️",
            "error": "❌"
        }
        
        formatted_message = f"[{timestamp}] {icons.get(level, 'ℹ️')} {message}\n"
        
        # Insérer le message
        self.log_text.insert(tk.END, formatted_message)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def start_merge(self):
        if not self.input_folder.get():
            messagebox.showerror("Erreur", "Veuillez sélectionner un dossier d'entrée")
            return
            
        if not self.output_file.get():
            messagebox.showerror("Erreur", "Veuillez spécifier un fichier de sortie")
            return
            
        # Démarrer la fusion dans un thread séparé pour éviter de bloquer l'interface
        self.merge_button.config(state='disabled')
        self.progress_var.set(0)
        
        thread = threading.Thread(target=self.merge_files)
        thread.daemon = True
        thread.start()
        
    def merge_files(self):
        try:
            input_path = Path(self.input_folder.get())
            output_path = Path(self.output_file.get())
            
            # Trouver tous les fichiers Excel
            excel_files = list(input_path.glob("*.xlsx")) + list(input_path.glob("*.xls"))
            
            if not excel_files:
                self.root.after(0, lambda: messagebox.showerror("Erreur", 
                    "Aucun fichier Excel trouvé dans le dossier sélectionné"))
                self.root.after(0, lambda: self.merge_button.config(state='normal'))
                return
                
            self.root.after(0, lambda: self.log_message(f"Trouvé {len(excel_files)} fichiers Excel", "info"))
            self.root.after(0, lambda: self.status_var.set(f"🔄 Fusion de {len(excel_files)} fichiers..."))
            
            all_dataframes = []
            total_files = len(excel_files)
            
            for i, file_path in enumerate(excel_files):
                try:
                    self.root.after(0, lambda f=file_path: self.log_message(f"Traitement de: {f.name}", "info"))
                    
                    # Lire le fichier Excel avec gestion d'erreurs améliorée
                    try:
                        df = pd.read_excel(file_path, engine='openpyxl')
                    except:
                        # Essayer avec xlrd pour les anciens fichiers .xls
                        df = pd.read_excel(file_path, engine='xlrd')
                    
                    # Vérifier que le DataFrame n'est pas vide
                    if df.empty:
                        self.root.after(0, lambda f=file_path: self.log_message(f"Fichier vide ignoré: {f.name}", "warning"))
                        continue
                    
                    # Nettoyer les noms de colonnes (supprimer les espaces en début/fin)
                    df.columns = df.columns.str.strip()
                    
                    # Ajouter une colonne avec le nom du fichier source si demandé
                    if self.add_source_column.get():
                        df['Fichier_Source'] = file_path.name
                    
                    self.root.after(0, lambda f=file_path, rows=len(df), cols=len(df.columns): 
                        self.log_message(f"✓ {f.name}: {rows} lignes, {cols} colonnes", "success"))
                    
                    all_dataframes.append(df)
                    
                    # Mettre à jour la progression
                    progress = (i + 1) / total_files * 100
                    self.root.after(0, lambda p=progress: self.progress_var.set(p))
                    
                except Exception as e:
                    self.root.after(0, lambda f=file_path, err=str(e): 
                        self.log_message(f"Erreur lors du traitement de {f.name}: {err}", "error"))
                    continue
            
            if not all_dataframes:
                self.root.after(0, lambda: messagebox.showerror("Erreur", 
                    "Aucun fichier n'a pu être lu correctement"))
                self.root.after(0, lambda: self.merge_button.config(state='normal'))
                return
            
            # Fusionner tous les DataFrames avec gestion des colonnes
            self.root.after(0, lambda: self.log_message("Fusion des données...", "info"))
            
            # Normaliser les colonnes avant la fusion
            self.root.after(0, lambda: self.log_message("Normalisation des colonnes...", "info"))
            normalized_dataframes = self.normalize_dataframes(all_dataframes)
            
            # Fusionner les DataFrames normalisés
            merged_df = pd.concat(normalized_dataframes, ignore_index=True)
            
            # Gérer les en-têtes si nécessaire
            if self.ignore_headers.get():
                # Garder seulement les en-têtes du premier fichier
                first_df = normalized_dataframes[0]
                merged_df.columns = first_df.columns
            
            # Sauvegarder le fichier fusionné
            self.root.after(0, lambda: self.log_message("Sauvegarde du fichier fusionné...", "info"))
            merged_df.to_excel(output_path, index=False)
            
            # Succès
            self.root.after(0, lambda: self.log_message("Fusion terminée avec succès!", "success"))
            self.root.after(0, lambda: self.log_message(f"Fichier sauvegardé: {output_path}", "success"))
            self.root.after(0, lambda: self.log_message(f"Total de lignes: {len(merged_df)}", "success"))
            self.root.after(0, lambda: self.log_message(f"Total de colonnes: {len(merged_df.columns)}", "success"))
            
            self.root.after(0, lambda: self.status_var.set("🎉 Fusion terminée avec succès!"))
            self.root.after(0, lambda: messagebox.showinfo("🎉 Succès", 
                f"Fusion terminée avec succès!\n\n"
                f"📁 Fichier sauvegardé: {output_path}\n"
                f"📊 Total de lignes: {len(merged_df)}\n"
                f"📋 Total de colonnes: {len(merged_df.columns)}"))
            
        except Exception as e:
            self.root.after(0, lambda: self.log_message(f"Erreur: {str(e)}", "error"))
            self.root.after(0, lambda: self.status_var.set("❌ Erreur lors de la fusion"))
            self.root.after(0, lambda: messagebox.showerror("❌ Erreur", f"Une erreur s'est produite:\n{str(e)}"))
            
        finally:
            self.root.after(0, lambda: self.merge_button.config(state='normal'))
    
    def normalize_dataframes(self, dataframes):
        """Normalise les DataFrames pour qu'ils aient les mêmes colonnes"""
        if not dataframes:
            return dataframes
        
        # Collecter toutes les colonnes uniques de tous les DataFrames
        all_columns = set()
        for df in dataframes:
            all_columns.update(df.columns)
        
        # Trier les colonnes pour un ordre cohérent
        all_columns = sorted(list(all_columns))
        
        self.root.after(0, lambda: self.log_message(f"Colonnes détectées: {len(all_columns)}", "info"))
        self.root.after(0, lambda: self.log_message(f"Colonnes: {', '.join(all_columns[:5])}{'...' if len(all_columns) > 5 else ''}", "info"))
        
        # Normaliser chaque DataFrame
        normalized_dfs = []
        for i, df in enumerate(dataframes):
            try:
                # Créer un nouveau DataFrame avec toutes les colonnes
                normalized_df = pd.DataFrame()
                
                for col in all_columns:
                    if col in df.columns:
                        # La colonne existe, copier les données
                        normalized_df[col] = df[col]
                    else:
                        # La colonne n'existe pas, remplir avec NaN
                        normalized_df[col] = pd.NA
                
                normalized_dfs.append(normalized_df)
                self.root.after(0, lambda idx=i: self.log_message(f"DataFrame {idx+1} normalisé: {len(df.columns)} → {len(all_columns)} colonnes", "info"))
                
            except Exception as e:
                self.root.after(0, lambda idx=i, err=str(e): self.log_message(f"Erreur normalisation DataFrame {idx+1}: {err}", "error"))
                # En cas d'erreur, utiliser le DataFrame original
                normalized_dfs.append(df)
        
        return normalized_dfs
    
    def on_closing(self):
        """Arrête l'animation et ferme la fenêtre"""
        self.animation_running = False
        self.root.destroy()

def main():
    root = tk.Tk()
    app = ExcelMergerApp(root)
    
    # Gérer la fermeture de la fenêtre
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    root.mainloop()

if __name__ == "__main__":
    main()
